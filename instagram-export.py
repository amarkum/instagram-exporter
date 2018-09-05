import os
import sys

try:
    import xlwt
except ImportError:
    print("Installing Data & Time Python Library.")
    os.system("sudo -H pip install xlwt")
    import xlwt
    from xlwt import Workbook

try:
    import datetime
except ImportError:
    print("Installing Data & Time Python Library.")
    os.system("sudo -H pip install datetime")
    import datetime

try:
    from google.cloud import vision
except ImportError:
    print("Installing Google API Python Library.")
    os.system("sudo -H pip install --ignore-installed --upgrade google-cloud-vision")
    from google.cloud import vision

import io
import re
import unicodedata
import urllib2
import datetime

from zipfile import ZipFile
from google.cloud import vision
from google.cloud.vision import types

# Declare a Excel Workbook
wb = Workbook()

# Add Sheet to Excel WorkBook
sheet1 = wb.add_sheet('Photo Sheet')

file_input = ""

if len(sys.argv) != 2:
    if len(sys.argv) > 2:
        print("Too many Input file passed.\nUsage: python script.py [filename] OR python script.py ")
        exit(0)
    else:
        csv_input = raw_input("No Input File given\nPlease enter filename:")
        file_input = csv_input
else:
    file_input = sys.argv[1]

currentTime = str(datetime.datetime.now())

# name of Excel file
excel_output_file = "RESPONSE_" + currentTime + ".xls"

try:
    os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = "key.json"
except:
    print("Please check if \"key.json\" file exists.")

# Instantiates a client

try:
    client = vision.ImageAnnotatorClient()
except:
    print("Can't Instantiate a Google Vision Client.")
    print("Please check your Internet Connection or check if Google API \"key.json\" is available.")
    exit(0)


def main():
    cleanup_images_jpg()
    get_info_from_api_flush_to_csv()
    create_zip_folder()
    cleanup_images_jpg()


# read all URLs from file
def get_info_from_api_flush_to_csv():
    try:
        text_file = open(file_input, 'r')
        file_text = text_file.read()
    except:
        print("Bad File. Can not Open")
        exit(0)
    text_file.close()

    urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\), ]|(?:%[0-9a-fA-F][0-9a-fA-F]))+', file_text)
    stop_word = '150x150'

    high_resolution_uri_list = []

    for uri in urls:
        if stop_word not in uri:
            high_resolution_uri_list.append(uri)

    count = 1
    for high_res_uri in high_resolution_uri_list:
        filename = str(count) + '.jpg'

        try:
            request = urllib2.Request(high_res_uri)
            img = urllib2.urlopen(request).read()
            with open(filename, 'w') as image:
                image.write(img)
            count = count + 1
            print("Downloaded/OK |" + " URL " + high_res_uri)
        except:
            print("Not Downloaded/Failed - Image doesn't Exists for the URL - " + high_res_uri)

    print("\nImage Download Complete.\n")

    photo_cell_row = 0

    for i in range(count - 1):
        filename = str(i + 1) + ".jpg"

        photo_name = "Photo" + str(i + 1)

        print(str("200 | OK | " + filename))

        # Get the Full Image Path
        image_name = os.path.join(
            os.path.dirname(__file__), filename)

        cell_count = 0

        # Loads the image into memory
        with io.open(image_name, 'rb') as image_file:
            content = image_file.read()
        image = types.Image(content=content)

        # LABEL_DETECTION

        response = client.label_detection(image=image)
        labels = response.label_annotations

        # Write Label(s) and Score Attribute  into Excel WorkBook

        lbl_count = 0
        for label in labels:
            lbl_count = lbl_count + 1
            cell_count = cell_count + 1

            sheet1.write(cell_count, photo_cell_row, "LABEL" + str(lbl_count))
            sheet1.write(cell_count, photo_cell_row + 1,
                         unicodedata.normalize('NFKD', label.description).encode('ascii', 'ignore'))
            cell_count = cell_count + 1
            sheet1.write(cell_count, photo_cell_row, "SCORE" + str(lbl_count))
            sheet1.write(cell_count, photo_cell_row + 1, float(label.score))

        # SAFE_SEARCH_DETECTION

        response = client.safe_search_detection(image=image)
        safe = response.safe_search_annotation

        # Names of likelihood from google.cloud.vision.enums
        likelihood_name = ('UNKNOWN', 'VERY_UNLIKELY', 'UNLIKELY', 'POSSIBLE',
                           'LIKELY', 'VERY_LIKELY')

        # Write Safe Search Attribute  into Excel WorkBook

        sheet1.write(cell_count + 1, photo_cell_row, 'ADULT')
        sheet1.write(cell_count + 1, photo_cell_row + 1, likelihood_name[safe.adult])

        sheet1.write(cell_count + 2, photo_cell_row, 'MEDICAL')
        sheet1.write(cell_count + 2, photo_cell_row + 1, likelihood_name[safe.medical])

        sheet1.write(cell_count + 3, photo_cell_row, 'SPOOFED')
        sheet1.write(cell_count + 3, photo_cell_row + 1, likelihood_name[safe.spoof])

        sheet1.write(cell_count + 4, photo_cell_row, 'VIOLENCE')
        sheet1.write(cell_count + 4, photo_cell_row + 1, likelihood_name[safe.violence])

        sheet1.write(cell_count + 5, photo_cell_row, 'RACY')
        sheet1.write(cell_count + 5, photo_cell_row + 1, likelihood_name[safe.racy])

        cell_count = cell_count + 5

        # END

        # WEB SEARCH DETECTION

        web_detection = client.web_detection(image=image).web_detection
        web_entities = web_detection.web_entities

        # Write Image Description and Scores into Excel WorkBook
        web_entities_lbl = 0
        for web_entity in web_entities:
            web_entities_lbl = web_entities_lbl + 1
            cell_count = cell_count + 1

            sheet1.write(cell_count, photo_cell_row, "DESCRIPTION" + str(web_entities_lbl))
            sheet1.write(cell_count, photo_cell_row + 1,
                         unicodedata.normalize('NFKD', web_entity.description).encode('ascii', 'ignore'))
            cell_count = cell_count + 1
            sheet1.write(cell_count, photo_cell_row, "SCORE" + str(web_entities_lbl))
            sheet1.write(cell_count, photo_cell_row + 1, float(web_entity.score))

        visually_similar_images = web_detection.visually_similar_images

        # Write Visually Similar image URLs into Excel WorkBook
        visually_similar_images_lbl = 0
        for visually_similar_image in visually_similar_images:
            visually_similar_images_lbl = visually_similar_images_lbl + 1
            cell_count = cell_count + 1
            sheet1.write(cell_count, photo_cell_row, "URL" + str(visually_similar_images_lbl))
            sheet1.write(cell_count, photo_cell_row + 1, visually_similar_image.url)

        # Write Best Guess Label(s) into Excel WorkBook
        best_guess_labels = web_detection.best_guess_labels
        best_guess_labels_lbl = 0
        for best_guess_label in best_guess_labels:
            best_guess_labels_lbl = best_guess_labels_lbl + 1
            cell_count = cell_count + 1

            sheet1.write(cell_count, photo_cell_row, "BEST_LABEL" + str(best_guess_labels_lbl))
            sheet1.write(cell_count, photo_cell_row + 1,
                         unicodedata.normalize('NFKD', best_guess_label.label).encode('ascii', 'ignore'))

        # END

        # COLOR PROPERTIES DETECTION

        response = client.image_properties(image=image)
        props = response.image_properties_annotation
        colors = props.dominant_colors.colors

        # Write Colour(s) into Excel WorkBook
        color_lbl = 0
        for color in colors:
            color_lbl = color_lbl + 1
            cell_count = cell_count + 1

            red = int(color.color.red)
            green = int(color.color.green)
            blue = int(color.color.blue)

            sheet1.write(cell_count, photo_cell_row, "COLOUR" + str(color_lbl))
            sheet1.write(cell_count, photo_cell_row + 1, 'R:' + str(red) + ' G:' + str(green) + ' B:' + str(blue))

        # END

        # LANGUAGE PROPERTIES DETECTION

        response = client.document_text_detection(image=image)
        document = response.full_text_annotation
        pages = document.pages

        # Write Language(s) into Excel WorkBook
        for page in pages:
            property_field = page.property
            lang_lbl = 0
            for language in property_field.detected_languages:
                lang_lbl = lang_lbl + 1
                cell_count = cell_count + 1
                sheet1.write(cell_count, photo_cell_row, "LANGUAGE" + str(lang_lbl))
                sheet1.write(cell_count, photo_cell_row + 1, language.language_code)
        # END

        # Save the photo name as Photo1, Photo2...Photo n in Excel Workbook
        sheet1.write_merge(0, 0, photo_cell_row, photo_cell_row + 1, photo_name)
        photo_cell_row = photo_cell_row + 2

    wb.save(excel_output_file)
    print("\nFlushed content into Excel WorkBook.\n>> " + excel_output_file)


def get_all_file_paths(directory):
    # initializing empty file paths list
    image_paths = []

    # crawling through directory and subdirectories

    for root, directories, files in os.walk(directory):

        for image in files:

            if ".jpg" in image:
                # join the two strings in order to form the full image_path.
                image_path = os.path.join(root, image)
                image_paths.append(image_path)

            if excel_output_file in image:
                # join the two strings in order to form the full image_path.
                image_path = os.path.join(root, image)
                image_paths.append(image_path)

    # returning all image paths
    return image_paths


def create_zip_folder():
    # path to folder which needs to be zipped
    directory = './'

    # calling function to get all file paths in the directory
    file_paths = get_all_file_paths(directory)

    # writing files to a zipfile
    with ZipFile(str(currentTime) + '-IMGS.zip', 'w') as zip:
        # writing each file one by one
        for file in file_paths:
            zip.write(file)
    print('\nCompressed Zipped Folder Created.\n>> ' + str(currentTime) + '-IMGS.zip')


def cleanup_images_jpg():
    filelist = [f for f in os.listdir('./') if f.endswith(".jpg")]
    for f in filelist:
        os.remove(os.path.join('./', f))


if __name__ == "__main__":
    main()
