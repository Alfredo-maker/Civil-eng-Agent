import warnings
from dotenv import load_dotenv
from pydantic import BaseModel, Field
from langchain_openai import ChatOpenAI
from langchain_core.prompts import ChatPromptTemplate
from langchain_core.output_parsers import PydanticOutputParser
from langchain_classic.agents import AgentExecutor, create_tool_calling_agent
from tools import search_tool, analyze_beam_tool, save_tool, verify_ec2_tool, draw_live_cad_tool, generate_dxf_tool, draw_diagrams_tool, plot_matplotlib_tool


# Load environment variables (.env file)
load_dotenv()


# 2. Define Pydantic response schema with Field descriptions for better LLM adherence
class ResearchResponse(BaseModel):
    topic: str = Field(description="Main topic or question researched or analyzed")
    summary: str = Field(description="Comprehensive summary of research or calculation findings")
    sources: list[str] = Field(description="List of cited source titles, URLs, or calculation notes")
    tools_used: list[str] = Field(description="Names of tools used during the process")


# 3. Initialize model and parser
# API Key is automatically picked up from OPENAI_API_KEY in your .env file
llm = ChatOpenAI(
    model="gpt-4o-mini",
)

parser = PydanticOutputParser(pydantic_object=ResearchResponse)

# 4. Construct system prompt

prompt = ChatPromptTemplate.from_messages(
    [
        (
            "system",
            """
            You are an expert civil and structural engineering research assistant.

            Available Tools and Operational Guidelines:
            1. 'search_tool': Use for current, external, or factual real-world structural engineering references.
            2. 'analyze_beam_tool': Use for fundamental elastic beam calculations (max bending moments, shear forces, and peak stresses).
            3. 'verify_rc_beam': Use for Eurocode 2 (EN 1992-1-1) reinforced concrete verification, steel sizing (As_req vs As_prov), and capacity checks (M_Rd, V_Rd).
            4. 'draw_live_autocad': Use to draw the beam cross-section and elevation geometry directly in active AutoCAD.
            5. 'draw_diagrams_mv': Use to draw Shear Force Diagrams (SFD) and Bending Moment Diagrams (BMD) directly in active AutoCAD.
            6. 'plot_stresses_matplotlib': MANDATORY tool whenever 'plot', 'diagram', or 'graph' is mentioned. You MUST invoke this tool so a Matplotlib window opens on screen.
            7. 'save_text_to_file': MANDATORY when saving results. You MUST pass a full summary string or JSON object containing BOTH input parameters (Length, UDL, Section Dimensions) AND complete analysis/verification outputs (M_max, V_max, stresses, steel provided, EC2 utilization). Never pass inputs alone.

            Execution Guidelines:
            - Explicitly state structural assumptions (e.g., beam dimensions b x h, concrete/steel grades, support conditions) when invoking calculation and drawing tools.
            - Whenever drawing or plotting is requested, call BOTH AutoCAD tools ('draw_live_autocad', 'draw_diagrams_mv') AND 'plot_stresses_matplotlib' to ensure visual and CAD outputs are generated.
            - Keep explanations concise, professional, and clear.

            When your task is complete, your final response MUST strictly adhere to the following format instructions and contain no extra text:
            {format_instructions}
            """,
        ),
        ("placeholder", "{chat_history}"),
        ("human", "{query}"),
        ("placeholder", "{agent_scratchpad}"),
    ]
).partial(format_instructions=parser.get_format_instructions())

# 5. Initialize Agent & Executor with both tools
tools = [ search_tool,
        analyze_beam_tool,
        save_tool,
        verify_ec2_tool,
        draw_live_cad_tool,
        generate_dxf_tool,
        draw_diagrams_tool,
        plot_matplotlib_tool

]

agent = create_tool_calling_agent(llm=llm, prompt=prompt, tools=tools)
agent_executor = AgentExecutor(agent=agent, tools=tools, verbose=True)

# 6. Execute execution loop
if __name__ == "__main__":
    query = input("What can I help you research or calculate? ")

    raw_response = agent_executor.invoke({"query": query})

    # Extract raw output string
    raw_output = raw_response.get("output", "")

    # Sanitize markdown formatting if the model still wraps JSON in code fences
    cleaned_output = raw_output.strip().strip("```json").strip("```").strip()

    try:
        structured_response = parser.parse(cleaned_output)
        print("\n--- Parsed Agent Response ---")
        print(f"Topic: {structured_response.topic}\n")
        print(f"Summary: {structured_response.summary}\n")
        print(f"Sources: {structured_response.sources}\n")
        print(f"Tools Used: {structured_response.tools_used}")

    except Exception as e:
        print("\nError parsing response:", e)
        print("Raw Output Received:", raw_output)