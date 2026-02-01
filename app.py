import streamlit as st
from openai import OpenAI
import pandas as pd
from io import BytesIO

# Branding & UI
st.set_page_config(page_title="ZEN AI Structura", page_icon="🟢")
st.title("🟢 ZEN AI Structura")
st.markdown("### Intelligent Bill Processing")

# API Setup (Using Streamlit Secrets)
client = OpenAI(
    base_url="https://models.inference.ai.azure.com",
    api_key=st.secrets["GITHUB_TOKEN"]
)

uploaded_file = st.file_uploader("Upload Bill Image", type=["jpg", "png", "jpeg"])

if uploaded_file:
    st.image(uploaded_file, caption="Processing Target", use_container_width=True)
    
    if st.button("🚀 Generate Excel Report"):
        with st.spinner("AI is structured data..."):
            try:
                # Optimized prompt for Excel structure
                response = client.chat.completions.create(
                    messages=[{"role": "user", "content": "Extract: Date, Merchant, Item, Price. Return ONLY raw data separated by commas. No intro text."}],
                    model="gpt-4o",
                )
                
                # Logic to convert AI text to a real Excel file
                raw_data = response.choices[0].message.content
                rows = [line.split(',') for line in raw_data.strip().split('\n')]
                df = pd.DataFrame(rows, columns=['Date', 'Merchant', 'Item', 'Price'])
                
                # Memory buffer for the file
                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df.to_excel(writer, index=False)
                
                st.success("Success! Your structured report is ready.")
                st.download_button(
                    label="📥 Download .xlsx Report",
                    data=output.getvalue(),
                    file_name="ZEN_Bill_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"Logic Error: {e}")
