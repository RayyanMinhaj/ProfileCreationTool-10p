import streamlit as st
import requests

# Set the API URL
API_URL = "https://profile-creation-tool-10p-cqmi.vercel.app"

st.set_page_config(page_title="10Pearls User Profile Conversion Tool")

st.title("10Pearls User Profile Creation")
st.write('Converting old user profiles or CVs into a standardized format along with match percentage.')
st.markdown("""---""")


uploaded_file = st.file_uploader("Choose a file")
job_description = st.text_area("Job Description")

if st.button("Submit"):
    if uploaded_file is not None and job_description:
        # Send file and job description to Flask API
        response = requests.post(
            f"{API_URL}/upload",
            files={"file": uploaded_file},
            data={"job_description": job_description}
        )
        
        if response.status_code == 200:
            result = response.json()
            match_percentage = result.get("percentage_match")
            missing_keywords = result.get("missing_keywords")
            download_link = f"{API_URL}{result.get('download_link')}"

            st.write(f"Match Percentage: {match_percentage}")
            st.write(f"Missing Keywords: {missing_keywords}")

            # Provide download link
            st.markdown(f"[Download Converted File!]({download_link})")
        else:
            st.error(f"Error: {response.json().get('error')}")
    else:
        st.error("Please upload a file and provide a job description.")

st.markdown("""---""")
