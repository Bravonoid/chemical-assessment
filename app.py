import streamlit as st
import json
import datetime as dt
from docxtpl import DocxTemplate
from docx2pdf import convert
import io
import pandas as pd

import database

st.title("Chemical Assessment Audit App")

# General Information
st.header("Informasi Umum")

# Dropdown
divisions = [
    "Administration",
    "Central Services",
    "Concentrating",
    "Dewatering Plant",
    "Environmental",
    "Geo Engineering",
    "Learning & Organizational Development",
    "Maintenance Support",
    "Mill Optimization Construction",
    "MINE UNDERGROUND",
    "OCCUPATIONAL HEALTH & SAFETY",
    "OPERATIONS MAINTENANCE",
    "Operations Support",
    "PT Bamanat Amiete Papua",
    "PT Eksplorasi Nusa Jaya",
    "PT Major Drilling Indonesia",
    "PT Orica Mining Services",
    "PT Pangansari Utama Catering",
    "PT Prisma Kusuma Jaya",
    "PT Puncakjaya Power",
    "PT Redpath Indonesia",
    "PT RUC Cementation Indonesia",
    "SUPPLY CHAIN MANAGEMENT",
    "Surface Mine",
    "TRANSPORTATION & L/L FACILITY MANAGEMENT",
]

# 2 Columns
col1, col2 = st.columns(2)

with col1:
    division = st.selectbox("Divisi", divisions)
    section = st.text_input("Seksi")
    date = st.date_input("Hari & Tanggal Audit", "today", format="DD/MM/YYYY")
    auditor = st.text_input("Nama Auditor")
    companion = st.text_input("Pendamping")

with col2:
    department = st.text_input("Departemen")
    location = st.text_input("Lokasi")
    person_responsible = st.text_input("Nama Penanggung Jawab Area")
    chemical_coordinator = st.text_input("Nama Chemical Coordinator")

# Convert date to string
date = date.strftime("%d/%m/%Y")


metadata = {
    "division": division,
    "section": section,
    "date": date,
    "time_updated": dt.datetime.now().strftime("%H:%M:%S"),
    "auditor": auditor,
    "companion": companion,
    "department": department,
    "location": location,
    "person_responsible": person_responsible,
    "chemical_coordinator": chemical_coordinator,
}

# Line break
st.write("---")

questions = json.load(open("questions.json"))

st.header("Hasil Audit")

total_grade = 0
key_index = 0
index = 1

final_recap = []
for question in questions:
    st.subheader(f'{index}. {question["object"]}')
    total_question = 0
    question_grade = 0
    question_index = 1

    question_summary = []
    question_recap = {
        "object": f'{index}. {question["object"]}',
        "questions": [],
    }

    index += 1
    col1, col2 = st.columns(2)
    for subquestion in question["questions"]:
        answer = 0
        total_question += 1

        # Get the column
        col_index = question["questions"].index(subquestion) % 2
        if col_index == 0:
            col = col1
        else:
            col = col2

        with col:
            # Get the type
            if subquestion["type"] == "radio":
                # Get the options
                options = []
                grades = []
                for option in subquestion["options"]:
                    options.append(option["option"])
                    grades.append(option["grade"])
                answer = st.radio(subquestion["question"], options, key=key_index)

                # Add the grade
                question_grade += grades[options.index(answer)]
                key_index += 1

                if grades[options.index(answer)] <= 0:
                    question_summary.append(
                        f"{question_index}. {subquestion['question']}: {answer}"
                    )
                question_index += 1

                # Add to recap
                question_recap["questions"].append(
                    {
                        "question": subquestion["question"],
                        "answer": answer,
                        "grade": grades[options.index(answer)],
                    }
                )

            elif subquestion["type"] == "number":
                answer = st.number_input(
                    subquestion["question"], min_value=0, key=key_index
                )

                if answer >= subquestion["minimum"]:
                    question_grade += 1
                key_index += 1

                if answer < subquestion["minimum"]:
                    question_summary.append(
                        f"{question_index}. {subquestion['question']}: {answer}"
                    )
                question_index += 1

                # Add to recap
                question_recap["questions"].append(
                    {
                        "question": subquestion["question"],
                        "answer": answer,
                        "grade": 100 if answer >= subquestion["minimum"] else 0,
                    }
                )

            elif subquestion["type"] == "multiselect":
                options = []
                grades = []
                for option in subquestion["options"]:
                    options.append(option["option"])
                    grades.append(option["grade"])
                answer = st.multiselect(subquestion["question"], options, key=key_index)

                # Add the grade
                question_grade += sum([grades[options.index(q)] for q in answer])
                key_index += 1

                if sum([grades[options.index(q)] for q in answer]) <= 0:
                    question_summary.append(
                        f"{question_index}. {subquestion['question']}: {answer}"
                    )
                question_index += 1

                # Add to recap
                question_recap["questions"].append(
                    {
                        "question": subquestion["question"],
                        "answer": answer,
                        "grade": sum([grades[options.index(q)] for q in answer]),
                    }
                )

            # 2 Decimal places
            question_grade = round(question_grade, 2)

    # Adjust object with the question number
    object_name = f'{index-1}. {question["object"]}'

    # Average the grade
    average_grade = question_grade / total_question
    st.write(f"**Question Grade: {average_grade}**")
    total_grade += average_grade

    question_recap["grade"] = average_grade

    # Check if summary is empty
    if question_summary:
        question_recap["summary"] = question_summary
    else:
        question_recap["summary"] = ["-"]

    # Add to recap
    final_recap.append(question_recap)

# Finalize total grade
total_grade = total_grade / 14

# 2 Decimal places
total_grade = round(total_grade, 2)


status = ""
if total_grade >= 99:
    status = "Sangat Baik"
elif total_grade >= 75:
    status = "Diimplementasikan Baik"
elif total_grade >= 50:
    status = "Diimplementasikan Hanya Sebagian"
elif total_grade >= 25:
    status = "Implementasikan Lemah"
else:
    status = "Implementasi Buruk <25%"

metadata["total_grade"] = total_grade
metadata["status"] = status

# Compile the result
data = {
    "metadata": metadata,
    "recap": final_recap,
}

# Save data to json
with open("data.json", "w") as f:
    json.dump(data, f)

# Line break
st.write("---")

# Button to save to database
st.header("Recap")

# Show the summary
for recap in final_recap:
    # Check if it doesn't include "-"
    if "-" not in recap["summary"]:
        st.subheader(recap["object"])
        for summary in recap["summary"]:
            st.write(summary)


# Line break
st.write("---")

# Save to database and print to docx
st.header("Export to Docx")

col1, col2 = st.columns(2)

with col1:
    if st.button("Save to Database"):
        # Insert to database
        database.insert_update(data)
        st.success("Data has been saved to the database!")

# Download the docx
with col2:
    with open("report.docx", "rb") as f:
        # Print docx using template
        doc = DocxTemplate("template.docx")

        # Convert summary in final recap into strings with bullet points and remove the number
        for recap in final_recap:
            if "-" not in recap["summary"]:
                summary = ""
                for s in recap["summary"]:
                    s = s.split(". ")[1]
                    summary += f"- {s}\n"
                recap["summary"] = summary
            else:
                recap["summary"] = "-"

        context = {
            "division": division,
            "department": department,
            "section": section,
            "location": location,
            "date": date,
            "person_responsible": person_responsible,
            "auditor": auditor,
            "chemical_coordinator": chemical_coordinator,
            "companion": companion,
            "total_grade": total_grade,
            "status": status,
            "audits": final_recap,
        }

        doc.render(context)
        doc.save("report.docx")

        b = io.BytesIO(f.read())
        b.seek(0)
        st.download_button(
            label="Download Report",
            data=b,
            file_name="report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

# Line break
st.write("---")

# History
st.header("History")
# Get all documents from the collection
documents = database.get_all_documents(database.result)

# Check if there are documents
metadata = []
recaps = []
for document in documents:
    metadata.append(document["metadata"])
    recaps.append(document["recap"])

# Convert to dataframe
df = pd.DataFrame(metadata)
st.write(df)

st.write("Select a row to view the details.")
row = st.selectbox("Select a row", df.index)

# Show the details
st.write("Details:")
st.write(recaps[row])


# Show the total grade in 2 decimal places
st.sidebar.header(f"Total Grades: {total_grade:.2f}")
st.sidebar.write("Total Grade is calculated based on the answers given.")

# Navigation
st.sidebar.header("Navigation")
st.sidebar.write("1. [General Information](#informasi-umum)")
st.sidebar.write("2. [Audit Results](#hasil-audit)")
st.sidebar.write("3. [Recap](#recap)")
st.sidebar.write("4. [Save & Print](#save-print)")
st.sidebar.write("5. [History](#history)")
