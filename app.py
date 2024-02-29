import streamlit as st
import json
import datetime as dt
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Mm
import io
import pandas as pd
from PIL import Image

import database

st.title("Chemical Assessment Audit App")

# General Information
st.header("Informasi Umum")

# Dropdown
divisions = [
    {"name": "AVCO", "locations": ["GSE Ware House Avco"]},
    {
        "name": "Environmental",
        "locations": [
            "Enviro MP 21",
            "Nursery Check Poin",
        ],
    },
    {"name": "FTM", "locations": ["FM Portsite"]},
    {
        "name": "Geo Engineering",
        "locations": [
            "Bornite Office MP72 Ridge Camp",
            "ENJ Base Camp",
            "GBC Office",
            "Geology Surabaya Office",
            "MDI Shop Amole XC 3",
            "OB3 UG Geology",
            "UG Geotech DMLZ",
        ],
    },
    {
        "name": "Maintenance Support",
        "locations": [
            "KPI HD shop 34",
            "KPI Maintenance Shop",
            "KPI Marine-Porsite",
            "KPI Opeation 34 Chrusher",
            "KPI Shop 38",
            "PT. United Tractor",
        ],
    },
    {"name": "MIS", "locations": ["MIS Shop LL"]},
    {"name": "OD", "locations": ["Nemangkawi LIP"]},
    {
        "name": "Operation Maintenance",
        "locations": [
            "Amole Shop Besar MP74",
            "DMLZ Shop LFF",
            "DMLZ Shop UFF",
            "DOZ Shop",
            "GBC Locomotive Shop",
            "GBC Shop Crane Bay",
            "UFF 2590L DMLZ",
            "DMLZ Shop LFF Crane Bay Shop",
            "Hoist Electric (BG)",
            "Tambelo Shop (BG)",
        ],
    },
    {"name": "Operations Support", "locations": ["Fleet Operation (Batch Plant 34)"]},
    {"name": "Petrosea", "locations": ["Warehouse Petrosea 38"]},
    {"name": "PHMC", "locations": ["PHMC Lowland"]},
    {"name": "PJP", "locations": ["PJP Base Camp", "PJP Porsite"]},
    {"name": "PSU", "locations": ["PSU Maintenance Shop"]},
    {"name": "PT. Alas Emas Abadi", "locations": ["RC Clinic"]},
    {"name": "PT. Sandvik SMC", "locations": [" LIP - Lowland"]},
    {"name": "PT. SSB", "locations": ["LIP Lowland"]},
    {"name": "PT. Stamford", "locations": ["PT. Stamford - LIP"]},
    {"name": "PT. Trakindo Utama", "locations": ["LIP Lowland"]},
    {"name": "SCM", "locations": ["Grasberg", "Warehouse LIP", "Warehouse Porsite"]},
    {"name": "SOS", "locations": ["Clinic KK", "Clinic Portsite", "GBT Clinic"]},
    {
        "name": "Technical Services",
        "locations": ["DWP Porsite", "Pipe Line MP 34", "TRMP Construction Shop 34"],
    },
    {
        "name": "Underground Mine",
        "locations": [
            "UG CIP Office",
            "72 Batchplant",
            "RUC 2510 Shop BG",
            "Alimak Shop DMLZ",
            "Alimak Shop GBC",
            "Construction Storage GBC",
            "Electrical Redpath GBC",
            "MCM Shop GBC",
            "Raisebore Shop GBC",
        ],
    },
]

# 2 Columns
col1, col2 = st.columns(2)

with col1:
    division = st.selectbox("Divisi", [d["name"] for d in divisions])
    section = st.text_input("Seksi")
    date = st.date_input("Hari & Tanggal Audit", "today", format="DD/MM/YYYY")
    auditor = st.text_input("Nama Auditor")
    companion = st.text_input("Pendamping")

with col2:
    department = st.text_input("Departemen")
    location = st.selectbox(
        "Lokasi",
        divisions[[d["name"] for d in divisions].index(division)]["locations"],
    )
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
total_recap = 0
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
    question_evidences = []

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
                    total_recap += 1
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
                    question_grade += 100
                key_index += 1

                if answer < subquestion["minimum"]:
                    question_summary.append(
                        f"{question_index}. {subquestion['question']}: {answer}"
                    )
                    total_recap += 1
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
                    total_recap += 1
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

    # Add to recap
    question_recap["evidences"] = question_evidences

    # Adjust object with the question number
    object_name = f'{index-1}. {question["object"]}'

    # Average the grade
    average_grade = question_grade / total_question

    if average_grade >= 99:
        average_grade = 100

    st.write(f"**Question Grade: {average_grade}**")
    total_grade += average_grade

    # Add evidence via image upload if the grade is less than 100
    if average_grade < 100:
        uploaded_files = st.file_uploader(
            "Unggah Bukti Audit", accept_multiple_files=True, key=key_index
        )
        key_index += 1

        for uploaded_file in uploaded_files:
            if uploaded_file is not None:
                # Convert to image
                img = Image.open(uploaded_file)

                # Show the image
                st.image(img, caption=uploaded_file.name, use_column_width=True)

                image_bytes = io.BytesIO()
                img.save(image_bytes, format="PNG")

                # Add to recap
                question_evidences.append(
                    {
                        "name": uploaded_file.name,
                        "data": image_bytes.getvalue(),
                    }
                )

    question_recap["evidences"] = question_evidences
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
metadata["total_recap"] = total_recap

# Compile the result
data = {
    "metadata": metadata,
    "recap": final_recap,
}

# Line break
st.write("---")

# Button to save to database
st.header("Recap")
st.subheader(f"Total Recap: {total_recap}")

# Show the summary
for recap in final_recap:
    # Check if it doesn't include "-"
    if "-" not in recap["summary"]:
        st.subheader(recap["object"])
        for summary in recap["summary"]:
            st.write(summary)

        for evidence in recap["evidences"]:
            st.image(evidence["data"], caption=evidence["name"], use_column_width=True)


# Line break
st.write("---")

# Save to database and print to docx
st.header("Save & Print")

# Add recommendation link
st.subheader("Recommendation")
st.write(
    "Please put a link into the recommendation IMS document first, before saving to the database."
)
recommendation = st.text_input("Recommendation Link (IMS)")

# Add recommendation to the metadata
metadata["recommendation"] = recommendation

col1, col2 = st.columns([1, 3])

with col1:
    if st.button("Save to Database"):
        # Check if the recommendation is empty
        if recommendation == "":
            st.error("Please fill in the recommendation link.")
            st.stop()

        # Insert to database
        database.insert_update(data)
        st.success("Data has been saved to the database!")

# Download the docx
with col2:
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
            recap["summary"] = ""

        images = []
        for evidence in recap["evidences"]:
            # Save the image
            img = InlineImage(doc, io.BytesIO(evidence["data"]), width=Mm(40))
            images.append(img)

        recap["images"] = images

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

    with open("report.docx", "rb") as f:

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

if not metadata:
    st.write("No history found.")
    st.stop()

# Convert to dataframe
df = pd.DataFrame(metadata)
st.write(df)

st.write("Select a row to view the report.")
row = st.selectbox("Select a row", df.index)

# Render the report
doc = DocxTemplate("template.docx")

# Convert summary in final recap into strings with bullet points and remove the number
for recap in recaps[row]:

    if "-" not in recap["summary"]:
        summary = ""
        for s in recap["summary"]:
            s = s.split(". ")[1]
            summary += f"- {s}\n"
        recap["summary"] = summary
    else:
        recap["summary"] = ""

    images = []
    for evidence in recap["evidences"]:
        # Save the image
        img = InlineImage(doc, io.BytesIO(evidence["data"]), width=Mm(40))
        images.append(img)

    recap["images"] = images

context = {
    "division": df["division"][row],
    "department": df["department"][row],
    "section": df["section"][row],
    "location": df["location"][row],
    "date": df["date"][row],
    "person_responsible": df["person_responsible"][row],
    "auditor": df["auditor"][row],
    "chemical_coordinator": df["chemical_coordinator"][row],
    "companion": df["companion"][row],
    "total_grade": df["total_grade"][row],
    "status": df["status"][row],
    "audits": recaps[row],
}

doc.render(context)
doc.save("history.docx")

with open("history.docx", "rb") as f:

    b = io.BytesIO(f.read())
    b.seek(0)
    st.download_button(
        label="Download History",
        data=b,
        file_name="history.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
