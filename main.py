import streamlit as st
import pandas as pd
import openai
import json
import os

# -----------------------
# Secure API Key Handling
# -----------------------
# First, try to retrieve the API key from Streamlit secrets.
if "openai" in st.secrets and "api_key" in st.secrets["openai"]:
    openai.api_key = st.secrets["openai"]["api_key"]
else:
    # Fallback to environment variable
    openai_api_key = os.getenv("OPENAI_API_KEY")
    if not openai_api_key:
        st.error("OpenAI API key not found. Please set it in .streamlit/secrets.toml or as an environment variable.")
        st.stop()
    openai.api_key = openai_api_key

# -----------------------
# Cache Excel Data
# -----------------------
@st.cache_data
def load_excel_data(filename):
    try:
        excel_file = pd.ExcelFile(filename)
        # Assume the header row is the first row (index 0)
        sheet1 = excel_file.parse(excel_file.sheet_names[0])
        sheet2 = excel_file.parse(excel_file.sheet_names[1])
        sheet3 = excel_file.parse(excel_file.sheet_names[2])
        return sheet1, sheet2, sheet3
    except Exception as e:
        st.error(f"Error reading the Excel file: {e}")
        return None, None, None

sheet1, sheet2, sheet3 = load_excel_data("Pharma_Strategy_Template_V1.xlsx")
if sheet1 is None or sheet2 is None or sheet3 is None:
    st.stop()

# Uncomment the next line to inspect the columns if needed
# st.write("Sheet1 Columns:", sheet1.columns.tolist())

# -----------------------
# Extract Options from Sheet1
# -----------------------
# Role options: from columns B–D (indexes 1 to 3)
role_options = list(sheet1.columns[1:4])
# Lifecycle options: from columns E–I (indexes 4 to 8)
lifecycle_options = list(sheet1.columns[4:9])
# Customer Journey options: from columns J–M (indexes 9 to 12)
journey_options = list(sheet1.columns[9:13])

# -----------------------
# Helper Functions
# -----------------------
def filter_strategic_imperatives(df, role, lifecycle, journey):
    """
    Filters the strategic imperatives (in column "Strategic Imperative")
    where the corresponding cells in the given role, lifecycle, and journey columns contain an "x".
    """
    # Check that the columns exist in the dataframe
    if role not in df.columns or lifecycle not in df.columns or journey not in df.columns:
        st.error("The Excel file’s columns do not match the expected names for filtering.")
        return []
    
    try:
        filtered = df[
            (df[role].astype(str).str.lower() == 'x') &
            (df[lifecycle].astype(str).str.lower() == 'x') &
            (df[journey].astype(str).str.lower() == 'x')
        ]
        return filtered["Strategic Imperative"].dropna().tolist()
    except Exception as e:
        st.error(f"Error filtering strategic imperatives: {e}")
        return []

def generate_ai_output(customized_result, selected_differentiators):
    """
    Uses the OpenAI API to generate a 2-3 sentence description, estimated cost, and timeframe.
    Returns a dictionary with keys: description, cost, and timeframe.
    """
    differentiators_text = ", ".join(selected_differentiators) if selected_differentiators else "None"
    prompt = f"""
You are an expert pharmaceutical marketing strategist.
Given the following strategy description: "{customized_result}"
and the selected product differentiators: "{differentiators_text}",
please provide a short 2-3 sentence description of the strategic recommendation.
Also, provide an estimated cost range in USD and an estimated timeframe in months for implementation.
Return the output as a JSON object with keys "description", "cost", and "timeframe".
    """
    try:
        with st.spinner("Generating AI output..."):
            response = openai.ChatCompletion.create(
                model="gpt-3.5-turbo",
                messages=[
                    {"role": "system", "content": "You are an expert pharmaceutical marketing strategist."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0.7,
            )
        content = response.choices[0].message.content.strip()
        try:
            output = json.loads(content)
        except json.JSONDecodeError:
            st.error("Error decoding AI response. Please try again.")
            output = {"description": "N/A", "cost": "N/A", "timeframe": "N/A"}
        return output
    except Exception as e:
        st.error(f"Error generating AI output: {e}")
        return {"description": "N/A", "cost": "N/A", "timeframe": "N/A"}

# -----------------------
# Build the Streamlit Interface
# -----------------------
st.title("Pharma Strategy Online Tool")

# Step 1: Criteria Selection
st.header("Step 1: Select Your Criteria")
role_selected = st.selectbox("Select your role", role_options)
lifecycle_selected = st.selectbox("Select the product lifecycle stage", lifecycle_options)
journey_selected = st.selectbox("Select the customer journey focus", journey_options)

# Step 2: Strategic Imperatives
st.header("Step 2: Select Strategic Imperatives")
strategic_options = filter_strategic_imperatives(sheet1, role_selected, lifecycle_selected, journey_selected)
if not strategic_options:
    st.warning("No strategic imperatives found for these selections. Please try different options.")
else:
    selected_strategics = st.multiselect("Select up to 3 Strategic Imperatives", options=strategic_options, max_selections=3)

# Step 3: Product Differentiators
st.header("Step 3: Select Product Differentiators")
if "Product Differentiators" not in sheet2.columns:
    st.error("Sheet2 must have a column named 'Product Differentiators'.")
    st.stop()

product_diff_options = sheet2["Product Differentiators"].dropna().unique().tolist()
selected_differentiators = st.multiselect("Select up to 3 Product Differentiators", options=product_diff_options, max_selections=3)

# Generate and Display Results
if st.button("Generate Strategy"):
    if not selected_strategics:
        st.error("Please select at least one strategic imperative.")
    else:
        st.header("Strategic Recommendations")
        # Ensure Sheet3 has the necessary columns
        if "Strategic Imperative" not in sheet3.columns or "Result" not in sheet3.columns:
            st.error("Sheet3 must have columns named 'Strategic Imperative' and 'Result'.")
            st.stop()
        
        # Filter results based on selected strategic imperatives from Sheet3
        results_df = sheet3[sheet3["Strategic Imperative"].isin(selected_strategics)]
        if results_df.empty:
            st.info("No results found for the selected strategic imperatives.")
        else:
            for idx, row in results_df.iterrows():
                base_result = row["Result"]
                # Append product differentiators if selected
                differentiators_text = ", ".join(selected_differentiators) if selected_differentiators else ""
                customized_result = base_result
                if differentiators_text:
                    customized_result += f" (Customized with: {differentiators_text})"
                # Get AI-generated details
                ai_output = generate_ai_output(customized_result, selected_differentiators)
                st.subheader(customized_result)
                st.write(ai_output.get("description", "No description available."))
                st.write(f"**Estimated Cost:** {ai_output.get('cost', 'N/A')}")
                st.write(f"**Estimated Timeframe:** {ai_output.get('timeframe', 'N/A')}")
