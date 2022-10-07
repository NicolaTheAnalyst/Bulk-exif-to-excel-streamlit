import sys
import pandas as pd
import os
from PIL import Image, ExifTags
import glob
import streamlit as st



def main(uploaded_pics):
    counter = 0
    gps_info = {}
    exif_dict = {}
    dataframes = {}
    if uploaded_pics is not None:  # se effettivamente c'è qualcosa
        for element in uploaded_pics:
            ext = get_ext(element.name)
            im = Image.open(element)
            im_exif = getexifmethod(ext, im)
            counter += 1
            # https://stackoverflow.com/questions/69694259/create-dataframe-variables-inside-for-loop-group-dataframes

            if im_exif is None:
                st.write(f"exif data not available for {element.name[:40]}")
            else:
                try:
                    for key, val in im_exif.items():
                        if key in ExifTags.TAGS:
                            exif_dict[ExifTags.TAGS[key]] = val
                    try:  # Check if there is GPS info available
                        for key, val in exif_dict['GPSInfo'].items():
                            if key in ExifTags.GPSTAGS:
                                gps_info[ExifTags.GPSTAGS[key]] = val
                                st.write(str(val))
                        exif_dict.update(gps_info)
                        exif_dict.pop("GPSInfo")  # delete double GPSInfo tag
                    except KeyError:
                        st.write(f"GPS Info not found for {element.name[:40]}")
                        pass  # crack on
                    df = pd.DataFrame(list(exif_dict.items()), columns=['Tags', 'Values'])
                    dataframes[
                        f'df{counter}'] = df  # https://stackoverflow.com/questions/69694259/create-dataframe-variables-inside-for-loop-group-dataframes
                    csv = convert_df(dataframes[f'df{counter}'])
                    st.download_button(
                        label=f"Download {element.name[:30]} data as CSV",  # [:30] prende i primi trenta caratteri
                        data=csv,
                        file_name=f'df{element.name[:30]}.csv',
                        mime='text/csv',
                    )

                    df_as_string = df.astype(
                        str)  # converte il df to string per evitare l'errore "Conversion failed for column Values with type object'"
                    st.write(df_as_string)
                except:
                    st.write("Ni dobro:", str(sys.exc_info()[0]), "occurred.")
                    st.write(str(sys.exc_info()[1:]))
        st.write(dataframes)


def getexifmethod(ext, im): #temporary workaround
    if ext == ".tif" or ext == ".tiff":
        im_exif = im.getexif()
    else:
        im_exif = im._getexif()
    return im_exif

def get_ext(filename):
    extensions = ["jpg", "jpeg", "jpe", "jif", "jfif", "jfi", "tif", "tiff", "riff", "jpg", "jpeg", "jpe", "jif", "jfif", "jfi", "tif", "tiff", "riff"] #compatible extensions, it's not elegant but it works
    splitted_filename = filename.split(".")
    if splitted_filename[-1] in extensions: # se l'estensione è tra quelle consentite
        ext = "." + splitted_filename[-1]  # prende solo l'estensione dal filename e ci mette il punto davanti
        return ext
    else:
        st.write("L'estensione non è tra quelle consentite")
        ext = "." + splitted_filename[-1]  # prende solo l'estensione dal filename e ci mette il punto davanti
        return ext # return it regardless

def convert_df(df): #converte il df in un csv
    return df.to_csv().encode('utf-8')

if __name__ == '__main__':
    st.header("Bulk exif to excel web app")
    filetypes = ["jpg", "jpeg", "jpe", "jif", "jfif", "jfi", "tif", "tiff", "riff", "jpg", "jpeg", "jpe", "jif", "jfif",
                 "jfi", "tif", "tiff", "riff"]
    uploaded_pics = st.sidebar.file_uploader("Carica le foto da cui esportare gli exif", type=filetypes,
                                     accept_multiple_files=True)

        #st.write(dataframes[f'df{counter}'].astype(str))

     #https://docs.streamlit.io/library/api-reference/widgets/st.file_uploader


    main(uploaded_pics)