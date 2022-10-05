import sys
import pandas as pd
import os
from PIL import Image, ExifTags
import glob
import streamlit as st



def main():
    counter = 0
    exif_dict = {}
    gps_info = {}
    ext = getext()
    path = getpath()
    filelist = glob.glob(path + "*" + ext)
    try:
        writer = pd.ExcelWriter('export.xlsx', engine='xlsxwriter')
    except PermissionError:
        print("Please close export.xlsx and try again.\nExiting..")
        sys.exit()
    for element in filelist:
        name = getname(counter, ext, path)
        worksheetname = name  #the name must not exceed the 31 chars
        im = Image.open(element)
        im_exif = getexifmethod(ext, im)
        counter += 1

        try:
            if im_exif is None:
                print("exif data not available")
            else:
                for key, val in im_exif.items():
                    if key in ExifTags.TAGS:
                        exif_dict[ExifTags.TAGS[key]] = val
                for key, val in exif_dict['GPSInfo'].items():
                    if key in ExifTags.GPSTAGS:
                        gps_info[ExifTags.GPSTAGS[key]] = val
                exif_dict.update(gps_info)
                exif_dict.pop("GPSInfo") #delete double GPSInfo tag

        except KeyError:
            print("Some exif tags might have been not found.")

        except:
            print("Ni dobro:", sys.exc_info()[0], "occurred.")
        df = pd.DataFrame(list(exif_dict.items()), columns=['Tags', 'Values'])
        df.to_excel(writer, index=False, sheet_name=worksheetname)
    writer.save()
    print("Exif tags exported successfully.")

def getname(count, ext, path):
    names = [os.path.basename(x) for x in glob.glob(path + "*" + ext)] #get the names from the file list
    n0 = str(names[count])
    if len(n0) > 31:
        n = n0[:31]
    else:
        n = n0
    return n

def getexifmethod(ext, im): #temporary workaround
    if ext == ".tif" or ext == ".tiff":
        im_exif = im.getexif()
    else:
        im_exif = im._getexif()
    return im_exif

def getpath():
    try:
        path = input(str("Provide the path where you stored your photos: "))
        if path[-1] == "\\":
            pass
        elif path[-1] != "\\":
            path = path + "\\"

    except:
        print("Ni dobro:", sys.exc_info()[0], "occurred.")

    if os.path.isdir(path):
        print(f"Path provided: {path}")
    else:
        print(f" '{path}' is not a valid path. The default path used is where this script is located.")
        path = os.path.abspath(os.path.dirname(__file__)) + "\\"  # takes the directory where the script is located
    return path

def getext():
    extensions = [".jpg", ".jpeg", ".jpe", ".jif", ".jfif", ".jfi", ".tif", ".tiff", ".riff", "jpg", "jpeg", "jpe", "jif", "jfif", "jfi", "tif", "tiff", "riff"] #compatible extensions, it's not elegant but it works
    ext = input("Provide the file extension of the photos ")
    if ext in extensions:
        if ext[0] == ".": #controlla che il primo carattere sia il punto eg. .jpeg
            print("")
        else:
            ext = "." + ext
    else:
        getext()
    return ext

def convert_df(df):
    # IMPORTANT: Cache the conversion to prevent computation on every rerun
    return df.to_csv().encode('utf-8')




if __name__ == '__main__':
    st.header("Bulk exif to excel web app")
    filetypes = ["jpg", "jpeg", "jpe", "jif", "jfif", "jfi", "tif", "tiff", "riff", "jpg", "jpeg", "jpe", "jif", "jfif",
                 "jfi", "tif", "tiff", "riff"]
    uploaded_pics = st.file_uploader("Carica le foto da cui esportare gli exif", type=filetypes,
                                     accept_multiple_files=True)
    counter = 0
    gps_info = {}
    exif_dict = {}
    dataframes = {}
    ext = ".jpg"
    if uploaded_pics is not None:  # se effettivamente c'è qualcosa
        for element in uploaded_pics:
            im = Image.open(element)
            im_exif = getexifmethod(ext, im)
            counter += 1
            #https://stackoverflow.com/questions/69694259/create-dataframe-variables-inside-for-loop-group-dataframes

            if im_exif is None:
                st.write("exif data not available")
            else:
                try:
                    for key, val in im_exif.items():
                        if key in ExifTags.TAGS:
                            exif_dict[ExifTags.TAGS[key]] = val
                    try: #Check if there is GPS info available
                        for key, val in exif_dict['GPSInfo'].items():
                            if key in ExifTags.GPSTAGS:
                                gps_info[ExifTags.GPSTAGS[key]] = val
                                st.write(str(val))
                        exif_dict.update(gps_info)
                        exif_dict.pop("GPSInfo")  # delete double GPSInfo tag
                    except KeyError:
                        st.write("GPS Info not found")
                        pass # crack on
                    df = pd.DataFrame(list(exif_dict.items()), columns=['Tags', 'Values'])
                    dataframes[f'df{counter}'] = df #https://stackoverflow.com/questions/69694259/create-dataframe-variables-inside-for-loop-group-dataframes
                    csv = convert_df(dataframes[f'df{counter}'])
                    st.download_button(
                        label=f"Download {element.name} data as CSV",
                        data=csv,
                        file_name=f'df{element.name}.csv',
                        mime='text/csv',
                    )

                    df_as_string = df.astype(
                        str)  # converte il df to string per evitare l'errore "Conversion failed for column Values with type object'"
                    st.write(df_as_string)
                except:
                    st.write("Ni dobro:", str(sys.exc_info()[0]), "occurred.")
                    st.write(str(sys.exc_info()[1:]))
        st.write(dataframes)
        st.write(dataframes[f'df{counter}'].astype(str))

     #https://docs.streamlit.io/library/api-reference/widgets/st.file_uploader


    pass
    #main()