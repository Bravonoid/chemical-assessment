from pymongo.mongo_client import MongoClient
import streamlit as st

uri = f'mongodb://{st.secrets["mongo"]["username"]}:{st.secrets["mongo"]["password"]}@ac-sizcvp9-shard-00-00.pugcxfs.mongodb.net:27017,ac-sizcvp9-shard-00-01.pugcxfs.mongodb.net:27017,ac-sizcvp9-shard-00-02.pugcxfs.mongodb.net:27017/?ssl=true&replicaSet=atlas-6bomex-shard-0&authSource=admin&retryWrites=true&w=majority'


@st.cache_resource
def init_connection():
    return MongoClient(uri)


client = init_connection()

# Send a ping to confirm a successful connection
try:
    client.admin.command("ping")
    print("Pinged your deployment. You successfully connected to MongoDB!")
except Exception as e:
    print(e)

# Get the database
db = client.ca

# Get the collection
result = db.result


def insert_update(document):
    result.insert_one(document)


# Get all documents from a collection
def get_all_documents(collection):
    result = collection.find()
    return result
