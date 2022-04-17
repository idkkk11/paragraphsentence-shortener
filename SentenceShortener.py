# Required libraries: Transformers, tensorflow, PyTorch, numpy, pandas.
from transformers import pipeline
import os
import pandas as pd
import numpy as np

# Reading of excel file and importing the data as an array.
# Reads from column C
size = 0 
df = pd.read_excel("spreadsheetfile.xlsx", index_col=None, usecols="C")
array = df.to_numpy()
flattened = array.flatten()
size = flattened.size

# Initialising the transformers summariser.
os.environ["CUDA_VISIBLE_DEVICES"] = "0"
summarizer = pipeline("summarization")
summarizer = pipeline("summarization", model="t5-base", tokenizer="t5-base", framework="tf")

# Loop to add (non)summarised data into a list.
list = ['Summarized Text']
for i in range (0, size):
    text = flattened[i]

    if len(text)>=250:      # Only summarises text with 250 characters or more, can be adjusted. (15 seconds)
        print("Summarizing line ", i, "...")
        summary_text = summarizer(text, max_length=100, min_length=5, do_sample=False)[0]['summary_text']
        list.append(summary_text)
    else:
        print("Line ", i, " not summarised.")
        list.append(text)

# Converts the list into a dataframe
df2 = pd.DataFrame(list)

# Converts and exports the dataframe into an excel file.
df2.to_excel('./thearray4.xlsx')