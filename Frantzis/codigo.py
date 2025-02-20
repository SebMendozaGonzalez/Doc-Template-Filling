import pandas as pd
from docx import Document
from typing import Dict
from docx2pdf import convert

def create_docs(df: pd.DataFrame) -> None:
    """This function revceives a dataframe and, for each row in the dataframe, fills up a word document template
        with the row's info and saves it in the "docs" folder

    Args:
        df (pandas dataframe): 
    """
    df.iloc[:, 3] = pd.to_datetime(df.iloc[:, 3], format='%m/%d/%Y')
    for i in range(len(df)):
        doc_copy = Document("./document_template.docx")
        
        patient_name: str = df.iloc[i,0]
        df_code: str = df.iloc[i,1]
        med_facility: str = df.iloc[i,2]
        date_of_loss = df.iloc[i, 3]
        billing_amount: str = df.iloc[i, 4]
        
        
        dic2: Dict[str, str] = {
            "Patient/Plaintiff: ": patient_name,
            "Date of Loss: ": date_of_loss.strftime('%m/%d/%Y') if pd.notna(date_of_loss) else False,
            "Balance Due: $": billing_amount
        }
        
        for holder, value in dic2.items():
            for para in doc_copy.paragraphs:
                if para.text.startswith(holder):
                    para.text = f"{holder}{value}" if value else ""
                elif para.text.startswith("This letter shall serve as notice that"):
                    para.text = f"This letter shall serve as notice that {med_facility} holds a balance with your client, {dic2['Patient/Plaintiff: ']}."
                    
                    
        doc_copy.save(f"./documentos/modified{i}.docx")
        
        convert(f"./documentos/modified{i}.docx", f"./pdfs/modified{i}.pdf")


if __name__ == "__main__":
    df: pd.DataFrame = pd.read_excel("./info.xlsx")

    create_docs(df)