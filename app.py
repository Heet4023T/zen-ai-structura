import streamlit as st
import os
from openai import OpenAI
import pandas as pd
from io import BytesIO

# 1. Page Config (Your Branding)
st.set_page_config(page_title="ZEN AI Structura", page_icon="🟢")
st.title("🟢 ZEN AI Structura")
st.info("Upload your bill image to generate a structured Excel report.")

# 2. Secure API Connection
# Make sure you added GITHUB_TOKEN in Streamlit Settings > Secrets!
client = OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=st.secrets["GITHUB_TOKEN"]
)

# 3. File Upload UI
uploaded_file = st.file_uploader("Choose a bill image (JPG/PNG)", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, caption="Target Bill", use_container_width=True)
    
    if st.button("🚀 Process with GitHub Models"):
        with st.spinner("AI is analyzing..."):
            try:
                # 4. AI Processing
                response = client.chat.completions.create(
                    messages=[{"role": "user", "content": "Extract Date, Item Name, and Total Price from this bill. Format as a CSV style list."}],
                    model="gpt-4o",
                )
                ai_result = response.choices[0].message.content
                
                # 5. Conversion to Excel
                # We split the AI text into a table format
                rows = [line.split(',') for line in ai_result.strip().split('\n')]
                df = pd.DataFrame(rows)
                
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False, header=False)
                
                st.success("Analysis Complete!")
                
                # 6. The Download Button
                st.download_button(
                    label="📥 Download Excel Report",
                    data=output.getvalue(),
                    file_name="bill_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Error: {str(e)}")
