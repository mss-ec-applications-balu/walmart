import os
import pandas as pd
from azure.ai.formrecognizer import DocumentAnalysisClient
from azure.identity import DefaultAzureCredential
from azure.ai.openai import OpenAIClient
import random

# Azure Form Recognizer (Document Intelligence) endpoint (replace with your actual endpoint)
AZURE_FORM_RECOGNIZER_ENDPOINT = "https://<your-form-recognizer-resource>.cognitiveservices.azure.com/"

# Azure OpenAI endpoint (replace with your actual endpoint)
AZURE_OPENAI_ENDPOINT = "https://<your-openai-resource>.openai.azure.com/"

# Initialize Azure Document Intelligence client with Azure AD authentication
credential = DefaultAzureCredential()  # Handles secure authentication
document_analysis_client = DocumentAnalysisClient(
    endpoint=AZURE_FORM_RECOGNIZER_ENDPOINT, 
    credential=credential
)

# Initialize Azure OpenAI client
openai_client = OpenAIClient(
    endpoint=AZURE_OPENAI_ENDPOINT,
    credential=credential
)

# Function to perform OCR using Azure Document Intelligence and extract key-value fields
def extract_fields_from_pdf(pdf_path):
    with open(pdf_path, "rb") as f:
        poller = document_analysis_client.begin_analyze_document("prebuilt-document", document=f)
        result = poller.result()
    
    # Extract key-value pairs (fields) from the document
    fields = {}
    for kv_pair in result.key_value_pairs:
        key = kv_pair.key.content if kv_pair.key else None
        value = kv_pair.value.content if kv_pair.value else None
        if key and value:
            fields[key] = value
    
    return fields

# Function to load prompts from a folder based on document type
def load_prompts_from_folder(prompt_folder, document_type):
    prompt_file_path = os.path.join(prompt_folder, f"{document_type}.txt")
    if os.path.exists(prompt_file_path):
        with open(prompt_file_path, "r") as f:
            prompts = f.readlines()
        return [prompt.strip() for prompt in prompts]
    else:
        return ["Summarize the content of the document."]  # Default prompt if no file is found

# Function to select a prompt for a document type (changed to take office into account)
def select_prompt_for_office_and_document_type(prompt_folder, office_name, document_type):
    # Add the office name to the prompt selection process
    office_prompt_folder = os.path.join(prompt_folder, office_name)  # Look for prompts specific to the office
    if not os.path.exists(office_prompt_folder):
        office_prompt_folder = prompt_folder  # Fall back to general prompts if office-specific folder doesn't exist
    
    prompts = load_prompts_from_folder(office_prompt_folder, document_type)
    return random.choice(prompts) if prompts else "Summarize the content of the document."

# Function to perform LLM processing using Azure OpenAI
def process_with_llm(text, prompt):
    response = openai_client.completions.create(
        engine="text-davinci-003",
        prompt=f"{prompt}\n\nText:\n{text}",
        max_tokens=1500,
        temperature=0.7
    )
    return response.choices[0].text.strip()

# Recursively find all PDFs in a directory and its subdirectories
def find_pdfs_in_folder(folder_path):
    pdf_files = []
    for root, dirs, files in os.walk(folder_path):
        for file in files:
            if file.endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
    return pdf_files

# Process PDFs inside folder structure (office/document type) and save results to Excel
def process_pdfs_to_excel_by_office(base_folder, prompt_folder, output_excel):
    # Initialize an Excel writer object
    writer = pd.ExcelWriter(output_excel, engine='openpyxl')
    
    # Initialize the list to collect all rows
    all_data = []

    # Traverse the office directories
    for office_name in os.listdir(base_folder):
        office_path = os.path.join(base_folder, office_name)
        
        if os.path.isdir(office_path):
            # Traverse the document types within each office
            for document_type in os.listdir(office_path):
                document_type_path = os.path.join(office_path, document_type)
                
                if os.path.isdir(document_type_path):
                    # Find all PDFs in the document type folder
                    pdf_files = find_pdfs_in_folder(document_type_path)

                    for pdf_file in pdf_files:
                        print(f"Processing {pdf_file} in {office_name}/{document_type}...")
                        
                        # Step 1: Perform OCR and field extraction
                        fields = extract_fields_from_pdf(pdf_file)
                        
                        # Step 2: Select a prompt for the office and document type from the prompt folder
                        prompt = select_prompt_for_office_and_document_type(prompt_folder, office_name, document_type)
                        
                        # Step 3: Perform LLM analysis on the extracted text using the selected prompt
                        ocr_text = "\n".join([f"{key}: {value}" for key, value in fields.items()])
                        llm_result = process_with_llm(ocr_text, prompt)
                        
                        # Step 4: Prepare data in row format
                        row_data = {
                            "Office": office_name,
                            "Document Type": document_type,
                            "PDF File": os.path.basename(pdf_file),
                            "LLM Result": llm_result
                        }
                        
                        # Include all extracted fields as separate columns
                        row_data.update(fields)

                        # Collect the results for Excel in row format
                        all_data.append(row_data)

    # Convert the collected data to a DataFrame and save it to Excel
    df_all = pd.DataFrame(all_data)
    df_all.to_excel(writer, sheet_name="All Results", index=False)
    
    # Save the final Excel file
    writer.save()
    print(f"All results saved to {output_excel}")

# Define the base folder containing office subfolders, prompt folder, and output Excel path
base_folder = "/path/to/office/folder"  # Update this path to your folder containing the office directories
prompt_folder = "/path/to/prompts/folder"  # Update this to your folder containing document type prompts
output_excel = "output_office_results_with_prompts.xlsx"

# Run the processing
process_pdfs_to_excel_by_office(base_folder, prompt_folder, output_excel)