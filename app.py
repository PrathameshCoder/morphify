import io
from flask import Flask, jsonify, request, make_response
from flask import Flask, render_template
from flask import Flask, request, send_file, render_template
import pandas as pd
from pydub import AudioSegment
from PIL import Image
from io import BytesIO
import os
import tempfile
import subprocess
from flask import Flask
from flask import request,render_template,redirect,url_for,send_file
import os,sys
from pdf2docx import parse
from typing import Tuple
from tkinter import Tk,messagebox
from tkinter import _tkinter
import tempfile
import wave
import ffmpeg

UPLOADER_FOLDER='D:/Conversion/templates'
app=Flask(__name__)
app.config['UPLOADER_FOLDER']=UPLOADER_FOLDER
app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/jpgtopng")
def jpgtopng():
    return render_template("jpgtopng.html")

@app.route("/pngtojpg")
def pngtojpg():
    return render_template("pngtojpg.html")

@app.route("/webptopng")
def webptopng():
    return render_template("webptopng.html")

@app.route("/bmptopng")
def bmptopng():
    return render_template("bmptopng.html")

@app.route("/pngtopdf")
def pngtopdf():
    return render_template("pngtopdf.html")

@app.route('/wavtomp3', methods=['GET'])
def wavtomp3():
    return render_template('wavtomp3.html')

@app.route('/heictojpg', methods=['GET'])
def heictojpg():
    return render_template('heictojpg.html')

@app.route('/help')
def help():
    return render_template('help.html')


@app.route('/csvtoxlsx', methods=['GET'])
def csvtoxlsx():
    return render_template('csvtoxlsx.html')

@app.route('/m4atomp3', methods=['GET'])
def m4atomp3():
    return render_template('m4atomp3.html')


@app.route('/api/jpgtopng', methods=['POST'])
def jpg_to_png():
    # Check if image is present in the request
    if 'image' not in request.files:
        return 'No image found!', 400
    
    # Get the image from the request
    img = request.files['image']
    
    # Check if image is in JPG format
    if img.filename.split('.')[-1] != 'jpg':
        return 'Invalid file format! Please select a JPG image.', 400
    
    # Open image and convert it to PNG
    img = Image.open(img)
    img = img.convert('RGB')
    
    # Create a BytesIO object to store the PNG image
    png_output = io.BytesIO()
    img.save(png_output, format='PNG')
    png_output.seek(0)
    
    # Return the PNG image as attachment
    return send_file(png_output, mimetype='image/png', as_attachment=True, download_name='converted.png')


@app.route('/api/pngtojpg', methods=['POST'])
def png_to_jpg():
    # check if image file is present in request
    if 'image' not in request.files:
        return jsonify({'error': 'No image file present in request'})

    # read the image file from the request
    image_file = request.files['image']

    # check if the file is a PNG image
    if not image_file.filename.lower().endswith('.png'):
        return jsonify({'error': 'Invalid file format. Please select a PNG image.'})

    # convert the PNG image to JPG format
    image = Image.open(image_file)
    with BytesIO() as buffer:
        image.convert('RGB').save(buffer, format='JPEG')
        img_data = buffer.getvalue()

    # send the converted JPG image as a downloadable attachment
    response = make_response(img_data)
    response.headers['Content-Type'] = 'image/jpeg'
    response.headers['Content-Disposition'] = 'attachment; filename=image.jpg'
    return response



@app.route('/api/webptopng', methods=['GET','POST'])
def webp_to_png():
    # Check if image is present in the request
    if 'image' not in request.files:
        return 'No image found!', 400
    
    # Get the image from the request
    img = request.files['image']
    
    # Check if image is in WebP format
    if img.filename.split('.')[-1] != 'webp':
        return 'Invalid file format! Please select a WebP image.', 400
    
    # Open image and convert it to PNG
    img = Image.open(img)
    img = img.convert('RGBA')
    img.load()  # required for png transparency issue
    img = img.convert('RGB')
    
    # Create a BytesIO object to store the PNG image
    png_output = io.BytesIO()
    img.save(png_output, format='PNG')
    png_output.seek(0)
    
    # Return the PNG image as attachment
    return send_file(png_output, mimetype='image/png', as_attachment=True, download_name='converted.png')
    


@app.route('/api/pngtopdf', methods=['POST'])
def png_to_pdf():
    # Get the uploaded file from the request
    png_file = request.files['image']
    # Check if file is in png format
    if not png_file.filename.endswith('.png'):
        return 'Error: File format not supported. Please select a PNG file.'
    # Check if file size is less than 5MB
    if len(png_file.read()) > 5 * 1024 * 1024:
        return 'Error: File size is too large. Please select a file smaller than 5MB.'
    # Reset file pointer to the beginning of file
    png_file.seek(0)
    # Open image from file
    img = Image.open(png_file)
    # Create a buffer to save PDF file
    pdf_buffer = io.BytesIO()
    # Save the image as PDF to the buffer
    img.save(pdf_buffer, format='PDF')
    # Reset buffer pointer to the beginning of file
    pdf_buffer.seek(0)
    # Create a response object with PDF file as attachment
    response = make_response(pdf_buffer.read())
    response.headers['Content-Disposition'] = 'attachment; filename=converted.pdf'
    response.mimetype = 'application/pdf'
    # Return the response object
    return response


@app.route('/api/wavtomp3', methods=['POST'])
def wavtomp3_api():
    # Check if audio file is provided
    if 'file' not in request.files:
        return 'No file provided'
    
    # Get the audio file
    audio_file = request.files['audio']
    
    # Check if audio file is WAV format
    if audio_file.filename.split('.')[-1].lower() != 'wav':
        return 'Invalid file format. Please select a WAV file.'
    
    # Create a temporary directory to save the audio file
    temp_dir = tempfile.mkdtemp()
    audio_path = os.path.join(temp_dir, audio_file.filename)
    
    # Save the audio file to the temporary directory
    audio_file.save(audio_path)
    
    # Convert the audio file to MP3 format using ffmpeg
    output_path = os.path.join(temp_dir, 'output.mp3')
    subprocess.run(['ffmpeg', '-i', audio_path, '-codec:a', 'libmp3lame', '-qscale:a', '2', output_path])
    
    # Return the MP3 file as a downloadable attachment
    return send_file(output_path, download_name='output.mp3', as_attachment=True)



@app.route('/api/heictojpg', methods=['POST'])
def heic_to_jpg():
    # check if file is present in the request
    if 'file' not in request.files:
        return 'No file selected', 400

    file = request.files['file']

    # check if the file is a heic image
    if not file.filename.endswith('.heic'):
        return 'Please select a .heic file', 400

    try:
        # convert heic image to jpg format
        img = Image.open(file)
        img.convert('RGB').save('temp.jpg', 'JPEG')

        # send the jpg image as a downloadable attachment
        return send_file('temp.jpg', as_attachment=True, download_name='converted.jpg')

    except Exception as e:
        return str(e), 500


@app.route('/api/csvtoxlsx', methods=['POST'])
def csvtoxlsx_api():
    # check if file is present in request
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    # check if file is csv
    if not file.filename.endswith('.csv'):
        return 'Selected file is not CSV!'
    # read csv file
    csv_data = pd.read_csv(io.StringIO(file.stream.read().decode('utf-8')))
    # convert to excel file
    excel_file = io.BytesIO()
    writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
    csv_data.to_excel(writer, index=False, sheet_name='Sheet1')
    writer.save()
    excel_file.seek(0)
    # return excel file as downloadable attachment
    return send_file(excel_file, download_name='converted.xlsx', as_attachment=True)


@app.route('/api/m4atomp3', methods=['POST'])
def m4atomp3_api():
    # Check if audio file exists
    if 'audio' not in request.files:
        return 'No audio file found!', 400
    
    # Save audio file
    audio_file = request.files['audio']
    audio_file.save('audio.m4a')
    
    # Convert to MP3
    os.system('ffmpeg -i audio.m4a audio.mp3')
    
    # Return MP3 file as attachment
    return send_file('audio.mp3', as_attachment=True)

@app.route('/pdftodocx',methods=['GET','POST'])
def pdftodocx():
    if request.method=="POST":
        def convert_pdf2docx(input_file:str,output_file:str,pages:Tuple=None):
           if pages:
               pages = [int(i) for i in list(pages) if i.isnumeric()]

           result = parse(pdf_file=input_file,docx_with_path=output_file, pages=pages)
           summary = {
               "File": input_file, "Pages": str(pages), "Output File": output_file
            }

           print("\n".join("{}:{}".format(i, j) for i, j in summary.items()))
           return result
        file=request.files['filename']
        if file.filename!='':
           file.save(os.path.join(app.config['UPLOADER_FOLDER'],file.filename))
           input_file=file.filename
           output_file=r"hello.docx"
           convert_pdf2docx(input_file,output_file)
           doc=input_file.split(".")[0]+".docx"
           print(doc)
           lis=doc.replace(" ","=")
           return render_template("docx.html",variable=lis)
    return render_template("pdftodocx.html")


@app.route('/docx',methods=['GET','POST'])
def docx():
    if request.method=="POST":
        lis=request.form.get('filename',None)
        lis=lis.replace("="," ")
        return send_file(lis,as_attachment=True)
    return  render_template("pdftodocx.html")

if __name__ == "__main__":
    app.run(debug=True)
