from matplotlib import image
import streamlit as st
import cv2
import PIL
from PIL import Image
import keras_ocr
import pandas as pd
import numpy as np
import os

if not os.path.exists("tempDirOCFR"):
  os.mkdir("tempDirOCFR")

pipeline = keras_ocr.pipeline.Pipeline()
st.title("Image to Excel Converter")
st.write("Upload an image and we'll extract the data and create an Excel file for you to download!")

uploaded_file = st.file_uploader("Upload an image", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:

    with open(os.path.join("tempDirOCFR",uploaded_file.name),"wb") as f:
         f.write(uploaded_file.getbuffer())
    image=Image.open(f"tempDirOCFR/{uploaded_file.name}")
    image=np.array(image)
    
    prediction_groups = pipeline.recognize([image])
    
    df=pd.DataFrame(prediction_groups[0], columns=['text', 'bbox'])
    
    coordinates=np.array(df.bbox)
    coordinates=np.array(sorted(coordinates, key=lambda coordinates : (int(coordinates[0][1]),int(coordinates[0][0]))))
    df2=pd.DataFrame()
    s=pd.Series([x for x in coordinates])
    df2['bbox']=s

    partition_by_rows=[]
    row=[df2.values[0][0]]
    next_val=df2.values[1][0]
    count=len(df2)
    row_count=0
    duplicate=0
    while row_count<count-1:
      while ((row[-1][1][1])-10<=next_val[0][1]).all() and ((row[-1][1][1])+10>=next_val[0][1]).all():
        if duplicate==0:
          row.append(next_val)
        row_count+=1
        if row_count<count-1:
          next_val=df2.values[row_count+1][0]
          duplicate=0
        else:
          break
      partition_by_rows.append(row)
      if row_count<count-1:
        row=[next_val]
        duplicate=1

    sorted_partitions=[]
    to_sort=list(partition_by_rows)
    for row in to_sort:
      for i in range(1, len(row)):
          key = row[i]
          # Move elements of arr[0..i-1], that are
          # greater than key, to one position ahead
          # of their current position
          j = i-1
          while j >=0 and key[0][0]< row[j][0][0] :
                  row[j+1] = row[j]
                  j -= 1
          row[j+1] = key
      sorted_partitions.append(row)

    data=[]
    df_values=df.values
    for row in sorted_partitions:
      data_row=[]
      for item in row:
        for x in df_values:
          if (item==x[1]).all():
            data_row.append(x[0])
            break
      data.append(data_row)
      data_row=[]

    final_data=[]
    for row_data in zip(sorted_partitions, data):
      row_data_lst=[]
      count=0
      merge_count=0
      for item in zip(row_data[0],row_data[1]):
        count+=1
        if merge_count>0:
          merge_count-=1
          continue
        row_data_lst.append(item[1])    
        for item2 in zip(row_data[0][count:],row_data[1][count:]):
          if (item[0]!=item2[0]).all() and item[0][1][0]-15<=item2[0][0][0]<=item[0][1][0]+15 :
            row_data_lst[-1]+=" "+item2[1]
            merge_count+=1
      final_data.append(row_data_lst)

    final_df=pd.DataFrame(final_data[1:],columns=final_data[0])    
    
    final_df.to_excel("output.xlsx", index=False)
    
    st.write("Extracted Data:")
    # st.dataframe(df)
    # st.download_button(label="Download Excel File", data=final_df.to_excel("extracted_data"), file_name="extracted_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with open("output.xlsx", "rb") as template_file:
        template_byte = template_file.read()

    st.download_button(label="Download Excel File", data=template_byte, file_name="output.xlsx", mime='application/octet-stream')

