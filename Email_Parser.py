import extract_msg
import pytesseract
from PIL import Image
import io
import re
import PyPDF2
from pdf2image import convert_from_bytes
import ollama
from tempfile import mkdtemp
import os
import shutil
from pdf2image import convert_from_path



# Extract text from scanned document 
def extract_text_from_path(path):
    prompt = """This is a scanned document. 
    You must perform high level ocr on this document in order to extract the text it contains.
    The text extraction must be content and layout aware, since the document might contain tables
    Your output should be the ONLY the extracted text, without any extra comments."""

    response = ollama.chat(
        model='llama3.2-vision',
        messages=[{
            'role': 'user',
            'content': prompt,
            'images': [path]
        }]
    )     
       
    extracted_text = response['message']['content']    
    return extracted_text
    
    
def process_image_with_vision(attachment):
    # Create a temporary directory to store the attachments
    temp_dir = mkdtemp()   
    try:
        file_name = attachment.longFilename
        file_path = os.path.join(temp_dir, file_name)
        
        # Save the attachment to the temporary directory
        with open(file_path, 'wb') as f:
            f.write(attachment.data)

        # Process the file
        print(file_path)
        extracted_text = extract_text_from_path(file_path)

    finally:
        # Delete the temporary directory and all its contents
        shutil.rmtree(temp_dir)
        # print(f"Temporary directory deleted: {temp_dir}")    
    return extracted_text




def process_scanned_pdf(pdf_path):
    """
    Converts each page of a scanned PDF into an image, extracts text from each image,
    combines the text, and deletes the temporary images.
    """
    # Create a temporary directory to store the images
    temp_dir = mkdtemp()
    print(f"Temporary directory created: {temp_dir}")

    try:
        # Convert PDF pages to images and save them in the temporary directory
        images = convert_from_path(pdf_path, output_folder=temp_dir, fmt='jpg')

        # Initialize an empty string to hold the combined text
        combined_text = ""

        # Extract text from each image and combine the results
        for i, image in enumerate(images):
            image_path = os.path.join(temp_dir, f"page_{i+1}.jpg")
            image.save(image_path, 'JPG')

            # Extract text from the image
            extracted_text = extract_text_from_path(image_path)

            # Append the extracted text to the combined text
            combined_text += extracted_text + "\n"

        return combined_text

    finally:
        # Delete the temporary directory and all its contents
        shutil.rmtree(temp_dir)
        # print(f"Temporary directory deleted: {temp_dir}")



def extract_text_from_scanned_pdf(attachment):
    temp_dir = mkdtemp()
    
    try:
        file_name = attachment.longFilename
        if file_name.endswith('.pdf'):  # Ensure the attachment is a PDF
            pdf_path = os.path.join(temp_dir, file_name)
                
            # Save the PDF attachment temporarily
            with open(pdf_path, 'wb') as f:
                f.write(attachment.data)
                
            # Process the extracted PDF
            extracted_text = process_scanned_pdf(pdf_path)

    finally:
        # Clean up by deleting the temporary directory and all its contents
        shutil.rmtree(temp_dir)
        # print(f"Temporary directory deleted: {temp_dir}")
    
    
    
    return extracted_text  


def process_pdf_attachment(attachment):
    extracted_text = ''
    
    pdf_data = attachment.data
    pdf_reader = PyPDF2.PdfReader(io.BytesIO(pdf_data))
    
    # Loop through all the pages of the PDF
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        extracted_text += page.extract_text()
    
    
    print("trying to find text...")   
    if not extracted_text.strip():
        print("Pdf is a scanned document...")        
        extracted_text = extract_text_from_scanned_pdf(attachment)  
    else:
        print("The string contains visible characters.")  
    
    
    return extracted_text



def process_msg_file(file_path):
    """ This function will go through the .msg file, if this is
        a conversation we will cut the messages based on who is speaking.
        The function will return ->.the list of messages, who is sending/receiving
        an other list containing the text from files that are attached to this .msg message 
    """
    
    msg = extract_msg.Message(file_path)
    email_body = msg.body
    sender = msg.sender
    date = msg.date
    
    # Example of splitting based on quoted responses or common reply markers
    if "From:" in email_body:
        # responses = email_body.split("From:")
        responses = re.split(r'(?=From:)', email_body)
        current_response = responses[0]  # The current email content
        conversation_history = responses[1:]
    
    attachments = msg.attachments
    attachments_as_text = []
    for attachment in attachments:
        
        filename = attachment.longFilename.lower()
               

        # Check if the attachment is an image
        if filename.endswith(('png', 'jpg', 'jpeg')):
            print("Found image")
            text = process_image_with_vision(attachment)
            attachments_as_text.append(text)
            

        # Check if the attachment is a PDF
        elif filename.endswith('.pdf'):
            # process_pdf_attachment(attachment)
            print("Found PDF")
            text = process_pdf_attachment(attachment)
            attachments_as_text.append(text)
        
            
        else:
            print("unused document found")
    
    return sender, date, current_response, responses, attachments_as_text





# Defining main function
def main():
    # Check if the directory exists
    # Update this path to the directory containing .msg files
    msg_directory = "230054"
    # extract_text_from_msg_files(msg_directory)
    
    
    list_of_messages = []
    if not os.path.isdir(msg_directory):
        print(f"Directory '{msg_directory}' does not exist.")
        return

    # List all files in the directory
    files = [f for f in os.listdir(msg_directory) if f.endswith('.msg')]

    if not files:
        print("No .msg files found in the directory.")
        return

    for file in files:
        file_path = os.path.join(msg_directory, file)
        try:
            # Parse the .msg file
            sender, date, current_response, responses, attachments_as_text = process_msg_file(file_path)
            for message in responses:
                list_of_messages.append(message)
        except Exception as e:
            print(f"Error processing {file}: {e}")
            
    # Check for duplicates
    old_length = len(list_of_messages)
    # if len(list_of_messages) != len(set(list_of_messages)):
    #     print("Duplicates found!")

    # Remove duplicates
    list_of_messages = list(set(list_of_messages))
    
    i = 1
    print("Emails")
    for msg in list_of_messages:
        print(i,"->",msg)
        i += 1
        
    if old_length != len(list_of_messages):
        print("===========================================================")
        print("Duplicate messages found!")
        print("Old length is ",old_length," and new length is ", len(list_of_messages))
        print("===========================================================")
        
    i = 1
    print("Attachments:")
    for msg in attachments_as_text:
        print(i,"->",msg)
        i += 1
        
    
        


# Using the special variable 
# __name__
if __name__=="__main__":
    main()