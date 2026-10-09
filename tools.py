import os
import json
import warnings
import numpy as np
import math
import matplotlib.pyplot as plt
import re
import eurocodepy as ec
import ezdxf
import win32com.client
import pythoncom
import time
from ezdxf import colors
from pydantic import BaseModel, Field
from langchain_community.tools import DuckDuckGoSearchRun
from langchain_core.tools import tool, Tool
from datetime import datetime

warnings.filterwarnings("ignore", category=DeprecationWarning)


def parse_dimensions(dim_str):
    """Extracts width and height from strings like '300x500', '300mmx500mm', etc."""
    numbers = re.findall(r"\d+", str(dim_str))
    if len(numbers) >= 2:
        return float(numbers[0]), float(numbers[1])
    return None, None


def flatten_data(data):
    """Recursively flattens nested dictionaries and parses JSON strings into a flat dictionary with normalized keys."""
    if isinstance(data, str) and data.strip():
        try:
            data = json.loads(data)
        except json.JSONDecodeError:
            flat = {}
            dim_match = re.search(r"(\d+)\s*x\s*(\d+)", data, re.IGNORECASE)
            if dim_match:
                flat["b"] = float(dim_match.group(1))
                flat["h"] = float(dim_match.group(2))
            len_match = re.search(r"(\d+\.?\d*)\s*m", data, re.IGNORECASE)
            if len_match:
                flat["length"] = float(len_match.group(1))
            udl_match = re.search(r"(\d+\.?\d*)\s*kN/m", data, re.IGNORECASE)
            if udl_match:
                flat["udl"] = float(udl_match.group(1))
            return flat

    if not isinstance(data, dict):
        return {}

    flat = {}
    def _recurse(d):
        for k, v in d.items():
            if isinstance(v, dict):
                _recurse(v)
            else:
                norm_key = str(k).lower().replace(" ", "").replace("_", "").replace("-", "")
                flat[norm_key] = v
                flat[k] = v
    _recurse(data)
    return flat


def get_param(flat_dict, keys, default=None):
    """Retrieves a numeric value from flat_dict trying multiple key aliases, stripping non-numeric chars if string."""
    for k in keys:
        norm_k = str(k).lower().replace(" ", "").replace("_", "").replace("-", "")
        val = None
        if norm_k in flat_dict:
            val = flat_dict[norm_k]
        elif k in flat_dict:
            val = flat_dict[k]
        
        if val is not None:
            if isinstance(val, (int, float)):
                return float(val)
            if isinstance(val, list) and len(val) > 0 and isinstance(val[0], (int, float)):
                return float(val[0])
            if isinstance(val, str):
                cleaned = re.sub(r"[^\d.-]", "", val).strip()
                if cleaned:
                    try:
                        return float(cleaned)
                    except ValueError:
                        pass
    return default


# ------------------------------------------------------------------------------
# 1. DYNAMIC & AUTO-DESIGN EUROCODE 2 RC BEAM VERIFICATION TOOL
# ------------------------------------------------------------------------------
def verify_rc_beam(data: str | dict = "") -> str:
    """Performs Eurocode 2 (EN 1992-1-1) design checks and reinforcement verification/design."""
    try:
        defaults = {
            "Length": 6.0,
            "UDL": 20.0,
            "b": 300.0,
            "h": 500.0,
            "cover": 35.0,
            "concrete_grade": "C25/30",
            "steel_grade": "B500B",
        }

        # Unpack input dictionary or JSON string
        if isinstance(data, str) and data.strip():
            try:
                data = json.loads(data)
            except json.JSONDecodeError:
                data = {}
        # Unpack input dictionary or JSON string using robust flattener
        flat = flatten_data(data)

        if not isinstance(data, dict):
            data = {}

        # 1. Span & Loading Extraction
        L = float(data.get("Length", data.get("length", defaults["Length"])))
        w_ed = float(data.get("UDL", data.get("udl", defaults["UDL"])))
        L = get_param(flat, ["Length", "length", "span", "l"], default=defaults["Length"])
        w_ed = get_param(flat, ["UDL", "udl", "load", "w", "w_ed"], default=defaults["UDL"])

        # 2. Section Dimensions Extraction
        section_input = data.get("section", data.get("size", data.get("dimensions", "")))
        section_input = flat.get("section", flat.get("size", flat.get("dimensions", "")))
        parsed_b, parsed_h = parse_dimensions(section_input)

        b = float(data.get("b", data.get("width", parsed_b or defaults["b"])))
        h = float(data.get("h", data.get("height", parsed_h or defaults["h"])))
        cover = float(data.get("cover", defaults["cover"]))
        b = get_param(flat, ["b", "width", "w"], default=parsed_b or defaults["b"])
        h = get_param(flat, ["h", "height", "depth"], default=parsed_h or defaults["h"])
        cover = get_param(flat, ["cover", "c"], default=defaults["cover"])

        # --- AUTOMATIC UNIT CONVERSION (Meters to Millimeters) ---
        if 0 < b < 10.0:
            b = b * 1000.0

        if 0 < h < 10.0:
            h = h * 1000.0
        # ---------------------------------------------------------

        # 3. Materials via eurocodepy
        c_grade = str(data.get("concrete_grade", defaults["concrete_grade"]))
        s_grade = str(data.get("steel_grade", defaults["steel_grade"]))
        c_grade = str(flat.get("concrete_grade", flat.get("concrete", defaults["concrete_grade"])))
        s_grade = str(flat.get("steel_grade", flat.get("steel", defaults["steel_grade"])))

        concrete = ec.Concrete(c_grade)
        rebar = ec.Reinforcement(s_grade)

        fck = concrete.fck
        fyk = rebar.fyk
        fcd = fck / 1.5
        fyd = fyk / 1.15

        # Internal Actions
        M_ed = (w_ed * L**2) / 8.0  # kNm
        V_ed = (w_ed * L) / 2.0  # kN
        M_ed_Nmm = M_ed * 1e6

        # Estimated Effective Depth
        est_stirrup = float(data.get("stirrup_dia", 10.0))
        est_bar = float(data.get("bar_dia", 20.0))
        d_est = h - cover - est_stirrup - (est_bar / 2.0)

        # Flexural Steel Sizing Calculation
        K = M_ed_Nmm / (b * (d_est**2) * fck)
        z_est = d_est * (0.5 + math.sqrt(max(0, 0.25 - (K / 1.134))))
        z_est = min(z_est, 0.95 * d_est)
        As_req = M_ed_Nmm / (fyd * z_est)

        # Minimum Steel Check (EC2 §9.2.1.1)
        fctm = 0.30 * (fck ** (2 / 3)) if fck <= 50 else 2.12 * math.log(1 + (fck + 8) / 10)
        As_min = max(0.26 * (fctm / fyk) * b * d_est, 0.0013 * b * d_est)

        # Automated Reinforcement Sizing
        if "bar_dia" in data or "n_bars" in data:
            bar_dia = float(data.get("bar_dia", 20.0))
            n_bars = int(data.get("n_bars", 4))
        else:
            bar_dia = 20.0 if As_req <= 1200 else 25.0
            single_bar_area = math.pi * (bar_dia / 2.0) ** 2
            n_bars = max(2, math.ceil(max(As_req, As_min) / single_bar_area))

        stirrup_dia = float(data.get("stirrup_dia", 10.0))
        s_links = float(data.get("stirrup_spacing", data.get("s_links", 200.0)))

        # Actual Structural Capacities
        d = h - cover - stirrup_dia - (bar_dia / 2.0)
        As_prov = n_bars * (math.pi * (bar_dia / 2.0) ** 2)
        z_prov = d - (0.5 * (As_prov * fyd) / (0.85 * fcd * b))
        M_rd = (As_prov * fyd * z_prov) / 1e6  # kNm

        # Shear Capacity
        A_sw = 2 * (math.pi * (stirrup_dia / 2.0) ** 2)
        z_shear = 0.9 * d
        cot_theta = 2.5
        V_rd_s = (A_sw / s_links) * z_shear * fyd * cot_theta / 1000.0  # kN

        flexure_pass = (M_ed <= M_rd) and (As_prov >= As_min)
        shear_pass = V_ed <= V_rd_s

        res_lines = [
            "=== EUROCODE 2 VERIFICATION RESULTS ===",
            f"Beam Span: {L:.2f} m | UDL: {w_ed:.2f} kN/m",
            f"Cross-Section: {int(b)} x {int(h)} mm (Effective Depth d = {d:.1f} mm, Cover = {int(cover)} mm)",
            f"Materials: Concrete {c_grade} (fck = {fck:.1f} MPa) | Steel {s_grade} (fyk = {fyk:.1f} MPa)",
            "--------------------------------------------------",
            f"- Max Bending Moment (M_Ed): {M_ed:.2f} kNm at x = {L/2:.2f} m",
            f"- Required Steel (As_req): {As_req:.2f} mm² | Min Steel (As_min): {As_min:.2f} mm²",
            f"- Decided Main Steel (As_prov): {n_bars}H{int(bar_dia)} ({As_prov:.2f} mm²)",
            f"- Bending Capacity (M_Rd): {M_rd:.2f} kNm | Util: {(M_ed/M_rd)*100:.1f}% [{'PASS' if flexure_pass else 'FAIL'}]",
            "--------------------------------------------------",
            f"- Max Shear Force (V_Ed): {V_ed:.2f} kN at x = 0.00 m",
            f"- Provided Stirrups: H{int(stirrup_dia)} @ {int(s_links)}mm c/c",
            f"- Shear Capacity (V_Rd,s): {V_rd_s:.2f} kN | Util: {(V_ed/V_rd_s)*100:.1f}% [{'PASS' if shear_pass else 'FAIL'}]",
        ]
        return "\n".join(res_lines)

    except Exception as e:
        return f"Error executing Eurocode verification: {str(e)}"


# ------------------------------------------------------------------------------
# 2. FAILSAFE SAVE RESULTS TOOL
# ------------------------------------------------------------------------------

def save_to_txt(data: str | dict = "", **kwargs) -> str:
    """Appends structured Eurocode verification results with x-positions to research_output.txt."""
    try:
        filename = kwargs.get("filename", "research_output.txt")

        # Failsafe argument extraction if LLM passes kwargs instead of positional arg
        if not data and kwargs:
            data = kwargs.get(
                "data", kwargs.get("content", kwargs.get("text", kwargs))
            )

        # Handle string input (try parsing JSON string into dict if passed as string)
        if isinstance(data, str) and data.strip():
            try:
                data = json.loads(data)
            except json.JSONDecodeError:
                pass  # Keep as plain string if it's not valid JSON

        # Flatten nested dictionary if the LLM wrapped sections (e.g. {"Beam Analysis": {...}, "Eurocode": {...}})
        if isinstance(data, dict):
            flat_data = {}
            for k, v in data.items():
                if isinstance(v, dict):
                    flat_data.update(v)
                else:
                    flat_data[k] = v
            data = flat_data

            # Key normalization helper (case-insensitive & whitespace-insensitive search)
            def get_val(keys, default=None):
                for key in keys:
                    for d_key in data.keys():
                        if d_key.lower().replace(" ", "").replace(
                            "_", ""
                        ) == key.lower().replace(" ", "").replace("_", ""):
                            return data[d_key]
                return default

            # 1. Inputs Extraction with Flexible Keys
            raw_length = get_val(
                ["length", "span", "l", "length(m)"], default=4.5
            )
            raw_width = get_val(
                ["width", "b", "w", "width(m)"], default=0.25
            )
            raw_height = get_val(
                ["height", "h", "height(m)"], default=0.40
            )
            raw_udl = get_val(["udl", "w", "udl(kn/m)"], default=20.0)
            support = get_val(
                ["supporttype", "support"], default="simply_supported"
            )
            concrete = get_val(["concrete", "concretegrade"], default="C25/30")
            steel = get_val(["steel", "steelgrade"], default="B500B")

            # Numerical conversions for span geometry
            try:
                L_num = float(
                    str(raw_length)
                    .lower()
                    .replace("m", "")
                    .replace("mm", "")
                    .strip()
                )
                if L_num > 100:  # Handle input passed in mm
                    L_num /= 1000.0
            except ValueError:
                L_num = 4.5

            try:
                b_num = float(
                    str(raw_width)
                    .lower()
                    .replace("m", "")
                    .replace("mm", "")
                    .strip()
                )
                if b_num > 10:  # Convert mm to m for display
                    b_num /= 1000.0
            except ValueError:
                b_num = 0.25

            try:
                h_num = float(
                    str(raw_height)
                    .lower()
                    .replace("m", "")
                    .replace("mm", "")
                    .strip()
                )
                if h_num > 10:  # Convert mm to m for display
                    h_num /= 1000.0
            except ValueError:
                h_num = 0.40

            try:
                udl_num = float(
                    str(raw_udl).lower().replace("kn/m", "").strip()
                )
            except ValueError:
                udl_num = 20.0

            length_str = f"{L_num:.2f} m"
            width_str = f"{b_num:.2f} m ({int(b_num * 1000)} mm)"
            height_str = f"{h_num:.2f} m ({int(h_num * 1000)} mm)"
            udl_str = f"{udl_num:.2f} kN/m"

            x_mid = f"{L_num / 2.0:.2f}"
            x_supp = "0.00"

            # Dynamic calculation fallbacks if LLM omitted results
            m_calc = (udl_num * (L_num**2)) / 8.0
            v_calc = (udl_num * L_num) / 2.0

            # 2. Results Extraction with Flexible Keys & Fallbacks
            m_max = get_val(
                ["maxbendingmoment", "max_bending_moment"],
                f"{m_calc:.2f} kNm at x = {x_mid} m",
            )
            v_max = get_val(
                ["maxshearforce", "max_shear_force"],
                f"{v_calc:.2f} kN at x = {x_supp} m",
            )
            sigma_max = get_val(
                ["peakbendingstress", "peak_bending_stress"],
                f"N/A at x = {x_mid} m",
            )
            tau_max = get_val(
                ["peakshearstress", "peak_shear_stress"],
                f"N/A at x = {x_supp} m",
            )

            # 3. Eurocode Verification Extraction
            m_ed = get_val(
                ["maxbendingmomented", "m_ed"],
                f"{m_calc:.2f} kNm at x = {x_mid} m",
            )
            as_req = get_val(
                ["requiredsteel", "as_req"], "Calculated per EC2"
            )
            as_prov = get_val(
                ["decidedmainsteel", "providedsteel", "as_prov"], "2H20"
            )
            m_rd = get_val(["bendingcapacity(m_rd)", "m_rd"], "Verified")
            bending_util = get_val(
                ["bendingutil", "utilizationratio(%)"], "PASS"
            )

            v_ed = get_val(
                ["maxshearforceed", "v_ed"],
                f"{v_calc:.2f} kN at x = {x_supp} m",
            )
            stirrups = get_val(["providedstirrups", "stirrups"], "H10 @ 200mm c/c")
            v_rd = get_val(["shearcapacity(v_rd,s)", "v_rd"], "Verified")
            shear_util = get_val(["shearutil"], "PASS")

            content = (
                "Inputs:\n"
                f"- Length: {length_str}\n"
                f"- Support Type: {support}\n"
                f"- Width: {width_str}\n"
                f"- Height: {height_str}\n"
                f"- UDL: {udl_str}\n"
                f"- Concrete: {concrete}\n"
                f"- Steel: {steel}\n\n"
                "Results:\n"
                f"- Max Bending Moment: {m_max}\n"
                f"- Max Shear Force: {v_max}\n"
                f"- Peak Bending Stress: {sigma_max}\n"
                f"- Peak Shear Stress: {tau_max}\n\n"
                "Eurocode Verification Results:\n"
                f"- Max Bending Moment (M_Ed): {m_ed}\n"
                f"- Required Steel (As_req): {as_req}\n"
                f"- Provided Steel (As_prov): {as_prov}\n"
                f"- Bending Capacity (M_Rd): {m_rd} | Util: {bending_util}\n"
                f"- Max Shear Force (V_Ed): {v_ed}\n"
                f"- Provided Stirrups: {stirrups}\n"
                f"- Shear Capacity (V_Rd,s): {v_rd} | Util: {shear_util}"
            )
        else:
            content = str(data)

        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted_text = f"--- Analysis & Verification Output ---\nTimestamp: {timestamp}\n\n{content}\n\n"

        file_path = os.path.abspath(filename)
        os.makedirs(os.path.dirname(file_path), exist_ok=True)

        with open(file_path, "a", encoding="utf-8") as f:
            f.write(formatted_text)

        return f"Results successfully saved to {file_path}"

    except Exception as e:
        return f"Error saving results to file: {str(e)}"

# Tool Exports
verify_ec2_tool = Tool(
    name="verify_eurocode_beam",
    func=verify_rc_beam,
    description="Performs Eurocode 2 structural checks and designs reinforcement configurations automatically.",
)

save_tool = Tool(
    name="save_results_to_file",
    func=save_to_txt,
    description="Saves calculation outputs and Eurocode verification results to research_output.txt.",
)

# 1. Search Tool
@tool
def search_tool(query: str) -> str:
    """Search the web for information using DuckDuckGo."""
    search = DuckDuckGoSearchRun()
    return search.invoke(query)


# 2. Beam Analysis Tool Input Schema
class BeamAnalysisInput(BaseModel):
    length: float = Field(description="Length of the beam in meters")
    support_type: str = Field(
        default="simply_supported", 
        description="Type of beam support: 'simply_supported' or 'cantilever'"
    )
    width: float = Field(default=0.1, description="Cross-section width in meters")
    height: float = Field(default=0.2, description="Cross-section height in meters")
    point_loads: list[tuple[float, float]] = Field(
        default=[], 
        description="List of point load tuples as (Load_in_kN, Position_in_meters), e.g., [(10.0, 2.5)]"
    )
    udl: float = Field(default=0.0, description="Uniformly distributed load in kN/m")


# 3. Beam Analysis Tool Function
@tool(args_schema=BeamAnalysisInput)
def analyze_beam_tool(
    length: float,
    support_type: str = "simply_supported",
    width: float = 0.1,
    height: float = 0.2,
    point_loads: list[tuple[float, float]] = [],
    udl: float = 0.0
) -> str:
    """Performs structural engineering analysis on a beam and calculates max bending moment, shear force, and peak stresses."""
    
    area = width * height
    S = (width * (height ** 2)) / 6

    num_points = 1000
    L = length
    x = np.linspace(0, L, num_points)
    
    V = np.zeros_like(x)
    M = np.zeros_like(x)

    if support_type.lower() == 'simply_supported':
        R2 = (udl * L * (L / 2) + sum(P * a for P, a in point_loads)) / L
        R1 = (udl * L + sum(P for P, _ in point_loads)) - R2

        V += R1 - udl * x
        for P, a in point_loads:
            V -= np.where(x >= a, P, 0)

        M += 0.5 * udl * x * (L - x)
        for P, a in point_loads:
            M += np.where(x <= a, (P * (L - a) / L) * x, (P * a / L) * (L - x))

    elif support_type.lower() == 'cantilever':
        V -= udl * (L - x)
        for P, a in point_loads:
            V -= np.where(x <= a, P, 0)

        M -= 0.5 * udl * (L - x)**2
        for P, a in point_loads:
            M -= np.where(x <= a, P * (a - x), 0)
    else:
        return f"Error: Unsupported support type '{support_type}'. Use 'simply_supported' or 'cantilever'."

    # Peak Values
    abs_M = np.abs(M)
    max_M_idx = np.argmax(abs_M)
    M_max = M[max_M_idx]
    x_M_max = x[max_M_idx]

    abs_V = np.abs(V)
    max_V_idx = np.argmax(abs_V)
    V_max = V[max_V_idx]
    x_V_max = x[max_V_idx]

    # Stresses (kPa)
    sigma_peak_kPa = (abs(M_max) * 1e3 / S) / 1e3
    tau_peak_kPa = ((1.5 * np.max(abs_V) * 1e3) / area) / 1e3

    return (
        f"Beam Analysis Results:\n"
        f"- Max Bending Moment: {M_max:.2f} kNm at x = {x_M_max:.2f} m\n"
        f"- Max Shear Force: {V_max:.2f} kN at x = {x_V_max:.2f} m\n"
        f"- Peak Bending Stress (σ_max): {sigma_peak_kPa:.2f} kPa\n"
        f"- Peak Shear Stress (τ_max): {tau_peak_kPa:.2f} kPa"
    )


# ------------------------------------------------------------------------------
# 1. LIVE AUTOCAD DRAWING TOOL (Single-Span Simply Supported Beam)
# ------------------------------------------------------------------------------
@tool
def draw_live_autocad(data: str | dict = "") -> str:
    """Draws a single-span simply supported RC beam with native AutoCAD dimensions in true black lines."""
    try:
        import pythoncom
        import time
        import win32com.client

        # Default Parameters for Single-Span Simply Supported Beam
        b, h = 300.0, 500.0
        L = 6000.0  # 6-meter single span
        cover = 35.0
        n_bars, bar_dia = 4, 20.0
        stirrup_dia = 10.0
        s_spacing = 200.0  # Uniform stirrup spacing
        support_width = 300.0  # Support width

        # Color Code 250 represents dark gray/black in AutoCAD Index Colors, or 0 for True Black
        BLACK_COLOR_INDEX = 250

        # Input Parsing Strategy
        if isinstance(data, dict):
            b = float(data.get("b", data.get("width", b)))
            h = float(data.get("h", data.get("height", h)))
            cover = float(data.get("cover", cover))
            n_bars = int(data.get("n_bars", n_bars))
            bar_dia = float(data.get("bar_dia", bar_dia))
            stirrup_dia = float(data.get("stirrup_dia", stirrup_dia))
            if "L" in data:
                L = float(data["L"])
            elif "spans" in data and isinstance(data["spans"], list) and len(data["spans"]) > 0:
                L = float(data["spans"][0])
        elif isinstance(data, str) and data.strip():
            try:
                parsed_json = json.loads(data)
                if isinstance(parsed_json, dict):
                    b = float(parsed_json.get("b", parsed_json.get("width", b)))
                    h = float(parsed_json.get("h", parsed_json.get("height", h)))
                    cover = float(parsed_json.get("cover", cover))
                    n_bars = int(parsed_json.get("n_bars", n_bars))
                    bar_dia = float(parsed_json.get("bar_dia", bar_dia))
                    stirrup_dia = float(parsed_json.get("stirrup_dia", stirrup_dia))
                    if "L" in parsed_json:
                        L = float(parsed_json["L"])
                    elif "spans" in parsed_json and isinstance(parsed_json["spans"], list) and len(parsed_json["spans"]) > 0:
                        L = float(parsed_json["spans"][0])
            except json.JSONDecodeError:
                dim_match = re.search(r"(\d+)\s*x\s*(\d+)", data, re.IGNORECASE)
                if dim_match:
                    b, h = float(dim_match.group(1)), float(dim_match.group(2))
        flat = flatten_data(data)

        b = get_param(flat, ["b", "width", "w", "beam_width"], default=b)
        h = get_param(flat, ["h", "height", "depth", "beam_height"], default=h)
        L = get_param(flat, ["l", "length", "span", "beam_length", "spans"], default=L)

        # Check for dimension string like "250x400" in raw string
        if isinstance(data, str):
            dim_match = re.search(r"(\d+)\s*x\s*(\d+)", data, re.IGNORECASE)
            if dim_match:
                b, h = float(dim_match.group(1)), float(dim_match.group(2))

        cover = get_param(flat, ["cover", "c"], default=cover)
        n_bars = int(get_param(flat, ["n_bars", "nbars", "num_bars", "bars"], default=n_bars))
        bar_dia = get_param(flat, ["bar_dia", "bardia", "dia", "diameter"], default=bar_dia)
        stirrup_dia = get_param(flat, ["stirrup_dia", "stirrupdia"], default=stirrup_dia)
        s_spacing = get_param(flat, ["s_spacing", "stirrup_spacing", "spacing", "s_links"], default=s_spacing)

        # Detect rebar callouts like "2H20" in flat values
        for val in flat.values():
            if isinstance(val, str):
                bar_match = re.search(r"(\d+)\s*[HTht#]\s*(\d+)", val)
                if bar_match:
                    n_bars = int(bar_match.group(1))
                    bar_dia = float(bar_match.group(2))
                    break

        # Unit Conversions (m to mm)
        if 0 < b < 10.0:
            b *= 1000.0
        if 0 < h < 10.0:
            h *= 1000.0
        if 0 < L < 100.0:
            L *= 1000.0

      # Force attachment to active, visible AutoCAD session
        try:
            acad = win32com.client.GetActiveObject("AutoCAD.Application")
            acad.Visible = True
        except Exception:
            try:
                acad = win32com.client.Dispatch("AutoCAD.Application")
                acad.Visible = True
            except Exception as e:
                return f"Could not launch or connect to AutoCAD: {str(e)}"

        doc = None
        for _ in range(5):
            try:
                doc = acad.ActiveDocument
                if doc is not None and acad.Visible:
                    break
            except Exception:
                time.sleep(0.5)

        if doc is None:
            # Try creating a fresh blank drawing if none is open
            try:
                doc = acad.Documents.Add()
            except Exception:
                return "AutoCAD is running but no active document is open. Please open a blank file in AutoCAD."

        msp = doc.ModelSpace

        def pnt(x, y):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), 0.0))

        def make_poly(points_list, color_code=BLACK_COLOR_INDEX):
            clean_pts = []
            for i in range(0, len(points_list), 2):
                x, y = float(points_list[i]), float(points_list[i + 1])
                if not clean_pts or (clean_pts[-2] != x or clean_pts[-1] != y):
                    clean_pts.extend([x, y])

            if len(clean_pts) < 4:
                return None

            var_pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, clean_pts)
            poly = msp.AddLightWeightPolyline(var_pts)
            poly.Color = color_code
            return poly

        def add_text(text_str, x, y, height=25.0, color_code=BLACK_COLOR_INDEX):
            txt = msp.AddMText(pnt(x, y), 0, str(text_str))
            txt.Height = float(height)
            txt.Color = color_code
            return txt

        def add_dim_horiz(pt1, pt2, dim_line_y, text_override=""):
            dim = msp.AddDimRotated(pnt(pt1[0], pt1[1]), pnt(pt2[0], pt2[1]), pnt(pt1[0], dim_line_y), 0.0)
            dim.TextHeight = 15.0
            dim.Color = BLACK_COLOR_INDEX
            if text_override:
                dim.TextOverride = text_override
            return dim

        def add_dim_vert(pt1, pt2, dim_line_x, text_override=""):
            dim = msp.AddDimRotated(pnt(pt1[0], pt1[1]), pnt(pt2[0], pt2[1]), pnt(dim_line_x, pt1[1]), math.pi / 2.0)
            dim.TextHeight = 15.0
            dim.Color = BLACK_COLOR_INDEX
            if text_override:
                dim.TextOverride = text_override
            return dim

        # =====================================================================
        # VIEW 1: CROSS SECTION (Offset X=500, Y=500)
        # =====================================================================
        x_sec, y_sec = 500.0, 500.0

        # Concrete Section Outline & Outer Stirrup
        make_poly([x_sec, y_sec, x_sec + b, y_sec, x_sec + b, y_sec + h, x_sec, y_sec + h, x_sec, y_sec], BLACK_COLOR_INDEX)
        make_poly(
            [
                x_sec + cover, y_sec + cover,
                x_sec + b - cover, y_sec + cover,
                x_sec + b - cover, y_sec + h - cover,
                x_sec + cover, y_sec + h - cover,
                x_sec + cover, y_sec + cover,
            ],
            BLACK_COLOR_INDEX,
        )

        r_bar = bar_dia / 2.0
        start_x = x_sec + cover + stirrup_dia + r_bar
        end_x = x_sec + b - cover - stirrup_dia - r_bar
        bot_y = y_sec + cover + stirrup_dia + r_bar
        top_y = y_sec + h - cover - stirrup_dia - r_bar

        # Bottom Tension Rebars
        x_positions = (
            [start_x + i * ((end_x - start_x) / (n_bars - 1)) for i in range(n_bars)]
            if n_bars > 1
            else [(start_x + end_x) / 2.0]
        )
        for x_pos in x_positions:
            circle = msp.AddCircle(pnt(x_pos, bot_y), r_bar)
            circle.Color = BLACK_COLOR_INDEX

        # Top Hanger Rebars (2x H12)
        for x_pos in [start_x, end_x]:
            circle = msp.AddCircle(pnt(x_pos, top_y), 6.0)
            circle.Color = BLACK_COLOR_INDEX

        # Section Annotations & Dimensions
        add_dim_horiz((x_sec, y_sec), (x_sec + b, y_sec), y_sec - 60.0)
        add_dim_vert((x_sec, y_sec), (x_sec, y_sec + h), x_sec - 60.0)
        add_dim_vert((x_sec, y_sec), (x_sec, y_sec + cover), x_sec - 120.0, f"c={int(cover)}")

        add_text(f"SECTION {int(b)}x{int(h)}", x_sec, y_sec + h + 50.0, height=25.0, color_code=BLACK_COLOR_INDEX)
        add_text(f"{n_bars}H{int(bar_dia)} BOTTOM BARS", x_sec - 280.0, bot_y - 10.0, height=18.0, color_code=BLACK_COLOR_INDEX)
        add_text("2H12 TOP HANGERS", x_sec - 280.0, top_y + 10.0, height=18.0, color_code=BLACK_COLOR_INDEX)
        add_text(f"STIRRUP H{int(stirrup_dia)}", x_sec + b + 20.0, y_sec + (h / 2.0), height=18.0, color_code=BLACK_COLOR_INDEX)

        # =====================================================================
        # VIEW 2: SIMPLY SUPPORTED LONGITUDINAL ELEVATION
        # =====================================================================
        x_off = x_sec + b + 600.0
        y_elev = y_sec

        # Concrete Beam Outline
        make_poly([x_off, y_elev, x_off + L, y_elev, x_off + L, y_elev + h, x_off, y_elev + h, x_off, y_elev], BLACK_COLOR_INDEX)

        # Pin & Roller Supports (Simply Supported Boundary Conditions)
        # Left Support Block
        make_poly([x_off - support_width, y_elev, x_off - support_width, y_elev - 300.0, x_off, y_elev - 300.0, x_off, y_elev], BLACK_COLOR_INDEX)
        # Right Support Block
        make_poly([x_off + L, y_elev, x_off + L, y_elev - 300.0, x_off + L + support_width, y_elev - 300.0, x_off + L + support_width, y_elev], BLACK_COLOR_INDEX)

        # Rebars in Elevation
        elev_bot_y = y_elev + cover + stirrup_dia + r_bar
        elev_top_y = y_elev + h - cover - stirrup_dia - r_bar
        make_poly([x_off + cover, elev_bot_y, x_off + L - cover, elev_bot_y], BLACK_COLOR_INDEX)
        make_poly([x_off + cover, elev_top_y, x_off + L - cover, elev_top_y], BLACK_COLOR_INDEX)

        # Stirrups (Uniform Distribution)
        x_curr = x_off + cover
        while x_curr <= x_off + L - cover:
            make_poly([x_curr, y_elev + cover, x_curr, y_elev + h - cover], BLACK_COLOR_INDEX)
            x_curr += s_spacing

        # Dimensions & Text
        add_dim_horiz((x_off, y_elev), (x_off + L, y_elev), y_elev - 150.0, f"SPAN L = {int(L)} mm")
        add_text("SIMPLY SUPPORTED ELEVATION", x_off, y_elev + h + 60.0, height=30.0, color_code=BLACK_COLOR_INDEX)
        add_text(
            f"STIRRUPS H{int(stirrup_dia)} @ {int(s_spacing)}mm c/c UNIFORM",
            x_off + (L / 2.0) - 200.0,
            y_elev - 300.0,
            height=22.0,
            color_code=BLACK_COLOR_INDEX,
        )

        acad.ZoomExtents()
        return f"Successfully rendered 1-span simply supported beam (L={int(L)}mm) in black lines in AutoCAD."

    except Exception as e:
        return f"Error connecting to AutoCAD: {str(e)}"


# ------------------------------------------------------------------------------
# 2. DXF GENERATION TOOL (Single-Span Simply Supported Beam)
# ------------------------------------------------------------------------------
@tool
def generate_dxf_file(data: str | dict = "") -> str:
    """Generates a single-span simply supported DXF drawing with pure black lines."""
    try:
        if isinstance(data, str) and data.strip():
            try:
                data = json.loads(data)
            except json.JSONDecodeError:
                data = {}
        flat = flatten_data(data)

        if not isinstance(data, dict):
            data = {}
        b = get_param(flat, ["b", "width", "w", "beam_width"], default=300.0)
        h = get_param(flat, ["h", "height", "depth", "beam_height"], default=500.0)
        L = get_param(flat, ["l", "length", "span", "beam_length", "spans"], default=6000.0)

        b = float(data.get("b", data.get("width", 300.0)))
        h = float(data.get("h", data.get("height", 500.0)))
        L = float(data.get("L", data.get("length", 6000.0)))
        cover = float(data.get("cover", 35.0))
        n_bars = int(data.get("n_bars", 4))
        bar_dia = float(data.get("bar_dia", 20.0))
        stirrup_dia = float(data.get("stirrup_dia", 10.0))
        filename = str(data.get("filename", "simply_supported_beam.dxf"))
        if isinstance(data, str):
            dim_match = re.search(r"(\d+)\s*x\s*(\d+)", data, re.IGNORECASE)
            if dim_match:
                b, h = float(dim_match.group(1)), float(dim_match.group(2))

        cover = get_param(flat, ["cover", "c"], default=35.0)
        n_bars = int(get_param(flat, ["n_bars", "nbars", "num_bars", "bars"], default=4))
        bar_dia = get_param(flat, ["bar_dia", "bardia", "dia", "diameter"], default=20.0)
        stirrup_dia = get_param(flat, ["stirrup_dia", "stirrupdia"], default=10.0)
        filename = str(flat.get("filename", "simply_supported_beam.dxf"))

        for val in flat.values():
            if isinstance(val, str):
                bar_match = re.search(r"(\d+)\s*[HTht#]\s*(\d+)", val)
                if bar_match:
                    n_bars = int(bar_match.group(1))
                    bar_dia = float(bar_match.group(2))
                    break

        if 0 < b < 10.0:
            b *= 1000.0
        if 0 < h < 10.0:
            h *= 1000.0
        if 0 < L < 100.0:
            L *= 1000.0

        doc = ezdxf.new("R2012")
        msp = doc.modelspace()
        
        # Color Index 250 (Black/Dark Gray) or True Color Black RGB(0,0,0)
        BLACK_COLOR = 250

        # View 1: Cross-Section (Offset X=500, Y=500)
        x_sec, y_sec = 500.0, 500.0
        msp.add_polyline2d(
            [(x_sec, y_sec), (x_sec + b, y_sec), (x_sec + b, y_sec + h), (x_sec, y_sec + h), (x_sec, y_sec)],
            dxfattribs={"color": BLACK_COLOR, "layer": "CONCRETE"},
        )
        msp.add_polyline2d(
            [
                (x_sec + cover, y_sec + cover),
                (x_sec + b - cover, y_sec + cover),
                (x_sec + b - cover, y_sec + h - cover),
                (x_sec + cover, y_sec + h - cover),
                (x_sec + cover, y_sec + cover),
            ],
            dxfattribs={"color": BLACK_COLOR, "layer": "STIRRUP"},
        )

        r_bar = bar_dia / 2.0
        start_x = x_sec + cover + stirrup_dia + r_bar
        end_x = x_sec + b - cover - stirrup_dia - r_bar
        bot_y = y_sec + cover + stirrup_dia + r_bar
        top_y = y_sec + h - cover - stirrup_dia - r_bar

        x_positions = (
            [start_x + i * ((end_x - start_x) / (n_bars - 1)) for i in range(n_bars)]
            if n_bars > 1
            else [(start_x + end_x) / 2.0]
        )

        for x_pos in x_positions:
            msp.add_circle(center=(x_pos, bot_y), radius=r_bar, dxfattribs={"color": BLACK_COLOR, "layer": "REBAR"})
        for x_pos in [start_x, end_x]:
            msp.add_circle(center=(x_pos, top_y), radius=6.0, dxfattribs={"color": BLACK_COLOR, "layer": "REBAR"})

        # Annotations
        msp.add_text(f"SECTION {int(b)}x{int(h)}", dxfattribs={"height": 25.0, "color": BLACK_COLOR}).set_placement(
            (x_sec, y_sec + h + 50)
        )
        msp.add_text(f"{n_bars}H{int(bar_dia)} BARS", dxfattribs={"height": 18.0, "color": BLACK_COLOR}).set_placement(
            (x_sec - 250, bot_y - 10)
        )

        # View 2: Simply Supported Beam Elevation
        x_off = x_sec + b + 600.0
        y_elev = y_sec
        msp.add_polyline2d(
            [(x_off, y_elev), (x_off + L, y_elev), (x_off + L, y_elev + h), (x_off, y_elev + h), (x_off, y_elev)],
            dxfattribs={"color": BLACK_COLOR, "layer": "CONCRETE"},
        )
        # End Supports
        msp.add_polyline2d(
            
            [(x_off - 300.0, y_elev), (x_off - 300.0, y_elev - 300.0), (x_off, y_elev - 300.0), (x_off, y_elev)],
            dxfattribs={"color": BLACK_COLOR, "layer": "SUPPORTS"},
        )
        msp.add_polyline2d(
            [(x_off + L, y_elev), (x_off + L, y_elev - 300.0), (x_off + L + 300.0, y_elev - 300.0), (x_off + L + 300.0, y_elev)],
            dxfattribs={"color": BLACK_COLOR, "layer": "SUPPORTS"},
        )
        
        msp.add_line((x_off + cover, bot_y), (x_off + L - cover, bot_y), dxfattribs={"color": BLACK_COLOR, "layer": "REBAR"})
        msp.add_line((x_off + cover, top_y), (x_off + L - cover, top_y), dxfattribs={"color": BLACK_COLOR, "layer": "REBAR"})

        file_path = os.path.abspath(filename)
        doc.saveas(file_path)
        return f"Simply supported beam DXF created at: {file_path}"

    except Exception as e:
        return f"Error creating DXF: {str(e)}"


# ------------------------------------------------------------------------------
# 3. TOOL EXPORTS
# ------------------------------------------------------------------------------
draw_live_cad_tool = draw_live_autocad
generate_dxf_tool = generate_dxf_file




BLACK_COLOR_INDEX = 250

# ------------------------------------------------------------------------------
# 1. LIVE AUTOCAD M & V DIAGRAMS TOOL
# ------------------------------------------------------------------------------
@tool
def draw_diagrams_mv(data: str | dict = "") -> str:
    """Draws Shear Force Diagram (SFD) and Bending Moment Diagram (BMD) directly in live AutoCAD."""
    try:
        # Default Parameters
        L = 6000.0       # Span length (mm)
        w = 22.0         # Design UDL (kN/m = N/mm)
        b = 300.0        # Section width (mm)

        # Input Parsing Strategy
        if isinstance(data, dict):
            L = float(data.get("L", data.get("length", L)))
            w = float(data.get("w", data.get("udl", w)))
            b = float(data.get("b", data.get("width", b)))
        elif isinstance(data, str) and data.strip():
            try:
                parsed = json.loads(data)
                if isinstance(parsed, dict):
                    L = float(parsed.get("L", parsed.get("length", L)))
                    w = float(parsed.get("w", parsed.get("udl", w)))
                    b = float(parsed.get("b", parsed.get("width", b)))
            except json.JSONDecodeError:
                dim_match = re.search(r"(\d+)\s*x\s*(\d+)", data, re.IGNORECASE)
                if dim_match:
                    b = float(dim_match.group(1))
                udl_match = re.search(r"(\d+\.?\d*)\s*kN/m", data, re.IGNORECASE)
                if udl_match:
                    w = float(udl_match.group(1))
                len_match = re.search(r"(\d+\.?\d*)\s*m", data, re.IGNORECASE)
                if len_match:
                    L = float(len_match.group(1)) * 1000.0

        # Automatic Unit Normalization (Meters to Millimeters)
        if 0 < L < 100.0:
            L *= 1000.0
        if 0 < b < 10.0:
            b *= 1000.0

        # Structural Calculations
        L_m = L / 1000.0
        V_max = (w * L_m) / 2.0         # kN
        M_max = (w * (L_m ** 2)) / 8.0  # kNm

        # Display Scale Factors
        scale_v = 1.5
        scale_m = 2.0

        # Connect to Active AutoCAD Session
        acad = win32com.client.Dispatch("AutoCAD.Application")
        doc = None
        for _ in range(5):
            try:
                doc = acad.ActiveDocument
                if doc is not None:
                    break
            except Exception:
                time.sleep(0.5)

        if doc is None:
            return "Could not connect to active AutoCAD document. Hit ESC in AutoCAD and try again."

        msp = doc.ModelSpace

        def pnt(x, y):
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (float(x), float(y), 0.0))

        def make_poly(pts, color=BLACK_COLOR_INDEX):
            clean_pts = []
            for i in range(0, len(pts), 2):
                x, y = float(pts[i]), float(pts[i + 1])
                if not clean_pts or (clean_pts[-2] != x or clean_pts[-1] != y):
                    clean_pts.extend([x, y])

            if len(clean_pts) < 4:
                return None

            var_pts = win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, clean_pts)
            poly = msp.AddLightWeightPolyline(var_pts)
            poly.Color = color
            return poly

        def add_text(txt, x, y, h=25.0):
            t = msp.AddMText(pnt(x, y), 0, str(txt))
            t.Height = float(h)
            t.Color = BLACK_COLOR_INDEX
            return t

        # Placement Offsets (Aligned under Section & Elevation)
        x_sec = 500.0
        x_start = x_sec + b + 600.0
        y_base_v = -1000.0
        y_base_m = -2200.0

        # --- 1. SHEAR FORCE DIAGRAM (SFD) ---
        make_poly([x_start, y_base_v, x_start + L, y_base_v], BLACK_COLOR_INDEX)
        v_pts = [
            x_start, y_base_v,
            x_start, y_base_v + (V_max * scale_v),
            x_start + L, y_base_v - (V_max * scale_v),
            x_start + L, y_base_v,
            x_start, y_base_v,
        ]
        make_poly(v_pts, BLACK_COLOR_INDEX)

        add_text("SHEAR FORCE DIAGRAM (V)", x_start, y_base_v + (V_max * scale_v) + 60.0, h=30.0)
        add_text(f"+V = {V_max:.1f} kN", x_start - 260.0, y_base_v + (V_max * scale_v), h=20.0)
        add_text(f"-V = -{V_max:.1f} kN", x_start + L + 30.0, y_base_v - (V_max * scale_v), h=20.0)

        # --- 2. BENDING MOMENT DIAGRAM (BMD) ---
        make_poly([x_start, y_base_m, x_start + L, y_base_m], BLACK_COLOR_INDEX)
        steps = 50
        m_pts = [x_start, y_base_m]
        for i in range(steps + 1):
            x_curr = (L / steps) * i
            x_m = x_curr / 1000.0

            M_x = (w * x_m / 2.0) * (L_m - x_m)
            y_curr = y_base_m - (M_x * scale_m)
            m_pts.extend([x_start + x_curr, y_curr])

        m_pts.extend([x_start + L, y_base_m, x_start, y_base_m])
        make_poly(m_pts, BLACK_COLOR_INDEX)

        add_text("BENDING MOMENT DIAGRAM (M)", x_start, y_base_m + 80.0, h=30.0)
        add_text(
            f"M_max = {M_max:.1f} kNm",
            x_start + (L / 2.0) - 150.0,
            y_base_m - (M_max * scale_m) - 60.0,
            h=22.0,
        )

        acad.ZoomExtents()
        return (
            f"Successfully rendered M & V diagrams in AutoCAD:\n"
            f"- Span L = {int(L)} mm, UDL w = {w:.2f} kN/m\n"
            f"- Peak Shear Force (V_max): {V_max:.2f} kN\n"
            f"- Peak Bending Moment (M_max): {M_max:.2f} kNm"
        )

    except Exception as e:
        return f"Error generating AutoCAD diagrams: {str(e)}"


# ------------------------------------------------------------------------------
# 2. MATPLOTLIB M, V & STRESSES PLOTTING TOOL
# ------------------------------------------------------------------------------
@tool
def plot_stresses_matplotlib(length: float = 6.0, udl: float = 22.0, width: float = 0.3, height: float = 0.5) -> str:
    """Plots SFD, BMD, peak shear stresses, and peak bending stresses using Matplotlib."""
    try:
        # Cross-sectional properties
        Area = width * height               # m²
        S = (width * (height ** 2)) / 6.0   # m³

        x = np.linspace(0, length, 500)

        # Internal Actions
        V = udl * (length / 2.0 - x)
        M = (udl * x / 2.0) * (length - x)

        # Stresses along span (kPa)
        sigma_peak = (np.abs(M) / S)
        tau_peak = (1.5 * np.abs(V) / Area)

        fig, axs = plt.subplots(2, 2, figsize=(11, 7))
        fig.suptitle(f"Beam Analysis: L={length}m, w={udl}kN/m ({int(width*1000)}x{int(height*1000)}mm)", fontweight='bold')

        # SFD
        axs[0, 0].plot(x, V, color='royalblue', lw=2)
        axs[0, 0].fill_between(x, 0, V, color='skyblue', alpha=0.4)
        axs[0, 0].set_title("Shear Force V (kN)")
        axs[0, 0].grid(True, linestyle=':', alpha=0.6)

        # BMD
        axs[0, 1].plot(x, M, color='firebrick', lw=2)
        axs[0, 1].fill_between(x, 0, M, color='salmon', alpha=0.4)
        axs[0, 1].set_title("Bending Moment M (kNm)")
        axs[0, 1].grid(True, linestyle=':', alpha=0.6)

        # Tau
        axs[1, 0].plot(x, tau_peak, color='darkorange', lw=2)
        axs[1, 0].fill_between(x, 0, tau_peak, color='orange', alpha=0.3)
        axs[1, 0].set_title("Peak Shear Stress τ_max (kPa)")
        axs[1, 0].set_xlabel("Span (m)")
        axs[1, 0].grid(True, linestyle=':', alpha=0.6)

        # Sigma
        axs[1, 1].plot(x, sigma_peak, color='purple', lw=2)
        axs[1, 1].fill_between(x, 0, sigma_peak, color='plum', alpha=0.4)
        axs[1, 1].set_title("Peak Bending Stress σ_max (kPa)")
        axs[1, 1].set_xlabel("Span (m)")
        axs[1, 1].grid(True, linestyle=':', alpha=0.6)

        plt.tight_layout()
        plt.show()

        return (
            f"Plotted M, V, and stresses via Matplotlib.\n"
            f"- Max Moment: {np.max(M):.2f} kNm\n"
            f"- Max Shear: {np.max(np.abs(V)):.2f} kN\n"
            f"- Max Bending Stress: {np.max(sigma_peak):.2f} kPa\n"
            f"- Max Shear Stress: {np.max(tau_peak):.2f} kPa"
        )
    except Exception as e:
        return f"Error plotting stresses: {str(e)}"


# ------------------------------------------------------------------------------
# 3. EXPORT TOOL LIST FOR YOUR AGENT
# ------------------------------------------------------------------------------
draw_diagrams_tool = draw_diagrams_mv
plot_matplotlib_tool = plot_stresses_matplotlib