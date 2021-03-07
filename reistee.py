import sys              # to work with command-line arguments
import zipfile          # to extract zips
import os               # various uses: navigation, deletion, paths 
import shutil           # mainly for moving and deleting files
import subprocess       # for running libreoffice in background
import functools        # used for rotating images by metadata
from fpdf import FPDF               # for creating pdf from images
from PIL import Image               # for opening images
from PyPDF2 import PdfFileMerger    # for merging pdfs




def remove_unnecassary_symbols(dirpath, filename):
    '''
    Removes "." "-" and "_" from filename and returns new filename

    Parameters
    ----------
    dirpath: Path of the directory that includes the file to rename

    filename: Filename of the file to rename

    '''

    filename = shutil.move(dirpath + "\\" + filename, dirpath + "\\" +
        (os.path.splitext(filename)[0]).replace(".","").replace("-","")
        .replace("_","") + os.path.splitext(filename)[1])

    return os.path.basename(filename)



def move_file_to_folder(file_with_path, curr_student):
    '''
    Moves a file to the main folder, and renames the file to the 
    reversed folder name, because student folders are named 
    "firstname lastname" and renamed files should be named 
    "lastname, fistname". Only used for files that dont have recognised
    fileendings.

    Parameters
    ----------
    file_with_path: File to move and rename, including the relative path

    curr_student: The stundent folder the files come from. Used for
    renaming the files accordingly

    '''

    folder_path = os.path.dirname(curr_student) + "\\"
    filename = os.path.basename(file_with_path)
    file_ext = os.path.splitext(filename)[1]

    folder_name_reversed = (
        (os.path.basename(curr_student)).split(" ")[1] + ", " + 
        (os.path.basename(curr_student)).split(" ")[0])

    #
    if os.path.isfile(folder_path + folder_name_reversed + file_ext):

        file_counter = 1

        while os.path.isfile(folder_path + folder_name_reversed + 
                                str(file_counter) + file_ext):
            file_counter += 1
            shutil.move(file_with_path, folder_path + 
                folder_name_reversed + str(file_counter) + file_ext)

        file_counter = 1

    else:
        shutil.move(file_with_path, folder_path + 
            folder_name_reversed + file_ext)




def categorise_file(dirpath, filename, file_lists):
    '''
    Categorise a given file into the correct file list. Currently 
    supportet are pictures, pdfs, docs, and txts. Returns false if
    file format is not recognized.

    Parameters
    ----------
    dirpath: The directory path the file is in

    filename: Filename of the file that shall be categorized

    file_lists: List of file-lists that each store the files of
    the correct format

    '''


    picture_files, pdf_files, doc_files, txt_files = file_lists
    file_to_conv_path = dirpath + "\\" + filename

    # Sort file into correct list. If file format is not recognized,
    # return false
    if (file_to_conv_path.endswith(".jpg") or 
        file_to_conv_path.endswith(".jpeg") or 
        file_to_conv_path.endswith(".png")):

        picture_files.append(file_to_conv_path)
    
    elif file_to_conv_path.endswith(".pdf"):

        pdf_files.append(file_to_conv_path)

    elif (file_to_conv_path.endswith(".doc") or 
            file_to_conv_path.endswith(".docx") or 
            file_to_conv_path.endswith(".odt") or 
            file_to_conv_path.endswith(".txt")):

        doc_files.append(file_to_conv_path)

    else:
        return False
        
    return True




# Thanks to Roman Odaisky for this solution to why the images were
# sometimes rotated, when all image viewers displayed it correctly.
# Problem is that when an image is taken sideways, the camera stores
# that images sideways and stores the rotation in metadata. Since 
# pillow removes Metadata, the image is "rotated" after opening.
# stackoverflow.com/questions/4228530/pil-thumbnail-is-rotating-my-image
def image_transpose_exif(im):
    """
    Apply Image.transpose to ensure 0th row of pixels is at the visual
    top of the image, and 0th column is the visual left-hand side.
    Return the original image if unable to determine the orientation.

    As per CIPA DC-008-2012, the orientation field contains an integer,
    1 through 8. Other values are reserved.

    Parameters
    ----------
    im: PIL.Image
    The image to be rotated.
    """

    exif_orientation_tag = 0x0112
    exif_transpose_sequences = [                   # Val  0th row  0th col
        [],                                        #  0    (reserved)
        [],                                        #  1   top      left
        [Image.FLIP_LEFT_RIGHT],                   #  2   top      right
        [Image.ROTATE_180],                        #  3   bottom   right
        [Image.FLIP_TOP_BOTTOM],                   #  4   bottom   left
        [Image.FLIP_LEFT_RIGHT, Image.ROTATE_90],  #  5   left     top
        [Image.ROTATE_270],                        #  6   right    top
        [Image.FLIP_TOP_BOTTOM, Image.ROTATE_90],  #  7   right    bottom
        [Image.ROTATE_90],                         #  8   left     bottom
    ]

    try:
        seq = exif_transpose_sequences[im._getexif()[exif_orientation_tag]]
    except Exception:
        return im
    else:
        return functools.reduce(type(im).transpose, seq, im)


def pic_to_pdf(pdf_filename, picture_files):
    '''
    Converts a list of pictures into one pdf-File that gets stored in 
    the directory given in pdf_filename. Generated pdfs contain _img_
    at the end of the filename for easier use elsewhere.

    Parameters
    ----------
    pdf_filename: Filename of combined pdf with all pictures. "_img_" 
    will be added to it.

    picture_files: List of Paths to pictures that shall be 
    combined into pdf.
    '''

    img_pdf_counter = 0
    pdf_list = []

    # create one pdf for every image, for scaling reasons 
    # with different sized images
    for picture_file in picture_files:

        # Open picture and rotate if stated in pictures metadata
        try:
            cover = Image.open(str(picture_file))
            cover_transposed = image_transpose_exif(cover)
            cover_transposed.save(str(picture_file))
            width, height = cover_transposed.size

            # create temporary pdf, with _imgpdf_ in filename
            pdf = FPDF(unit = "pt", format = [width, height])
            pdf.add_page()
            pdf.image(str(picture_file), 0, 0)

            img_pdf_counter += 1

            pdf.output(pdf_filename + "_imgpdf_" + 
                str(img_pdf_counter) + ".pdf", "F")

            pdf_list.append(pdf_filename + "_imgpdf_" + 
                str(img_pdf_counter) + ".pdf")
        except Image.UnidentifiedImageError:
            print("Can't open image: " + 
                os.path.basename(pdf_filename) + 
                "\\" + os.path.basename(picture_file) + " : " +
                "May be damaged or not correctly converted.")
            move_file_to_folder(picture_file, pdf_filename)


    # merge all _imgpdf_ files
    if len(pdf_list) > 0:
        img_merger = PdfFileMerger(strict=False)
        for pdf in pdf_list:
            img_merger.append(pdf)
        img_merger.write(pdf_filename + "_img_" + ".pdf")
        img_merger.close()

    # remove all _imgpdf_ Files
    for pdf in pdf_list:
        os.remove(pdf)


def check_libreoffice_install():
    '''
    Checks, if LibreOffice or OpenOffice is installed in the default
    directories or if it is listed in the path environmental
    variable. If not, returns an empty string
    '''

    libreoffice_x64_path = (
        "C:\\Program Files\\LibreOffice\\program\\soffice.exe")
    
    # I hope that's true
    libreoffice_x86_path = (
        "C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe")

    libreoffice_path = ""


    if(os.path.isfile(libreoffice_x86_path)):

        libreoffice_path = libreoffice_x86_path

    elif(os.path.isfile(libreoffice_x64_path)):
        libreoffice_path = libreoffice_x64_path

    elif(os.environ.get('Path').count("libreoffice\\program") > 0 or 
        os.environ.get('Path').count("openoffice\\program") > 0):

        libreoffice_path = "soffice"
    

    return libreoffice_path
    


def doc_to_pdf(pdf_filename, doc_files):
    '''
    Converts a list of odt, doc and docx Files into one pdf-File that 
    gets stored in the directory given in pdf_filename. Generated pdfs 
    contain _docs_ at the end of the filename for easier use elsewhere.
    A working installation of Libreoffice is requiered for running this
    method, and soffice.exe should be in PATH.

    Parameters
    ----------
    pdf_filename: Filename of combined pdf with all pictures. "_docs_" 
    will be added to it.

    doc_files: List of Paths to .odt, .doc or .docx Files that shall be 
    combined into pdf.
    '''

    base_dir = os.getcwd()
    doc_counter = 1
    converted_docs = []

    # navigate to every file and call libreoffice in that directory
    # one pdf for every file will be generated by libreoffice
    for doc_file in doc_files:   

        os.chdir(os.getcwd() + "\\" + os.path.dirname(doc_file))

        if not check_libreoffice_install() == "":


            libreoffice = subprocess.Popen(check_libreoffice_install() +
                " --convert-to "+ str(doc_counter) + 
                "py.pdf "+ "\"" + os.path.basename(doc_file)+ "\"")
            test = libreoffice.wait()
        
            converted_docs.append(os.path.dirname(doc_file) + "\\" + 
                os.path.splitext(os.path.basename(doc_file))[0] + "." + 
                str(doc_counter) + "py.pdf")

            doc_counter += 1


        # Should convert doc, docx and odt with MS Word. NOT TESTED! 
        else:
            new_name = (os.path.dirname(doc_file) + "\\" +
                os.path.splitext(os.path.basename(doc_file))[0] + 
                "." + str(doc_counter) + "py.pdf")

            try:
                word = os.client.DispatchEx("Word.Application")
                worddoc = word.Documents.Open(doc_file)
                worddoc.SaveAs(new_name, FileFormat = 17)
                converted_docs.append(os.path.dirname(doc_file) + "\\" + 
                    os.path.splitext(os.path.basename(doc_file))[0] + "." + 
                    str(doc_counter) + "py.pdf")
                doc_counter += 1
                worddoc.Close()
                word.Quit()

            except Exception as e:
                os.chdir(base_dir)
                print("No Office installed. Please install LibreOffice!")
                
                move_file_to_folder(doc_file, pdf_filename)

        os.chdir(base_dir)


    doc_counter = 1

    if not len(converted_docs) == 0:
    # merge all created pdfs into one pdf
        doc_merger = PdfFileMerger(strict=False)
        for converted_doc in converted_docs:
            doc_merger.append(converted_doc)

        doc_merger.write(pdf_filename + "_docs_" + ".pdf")
        doc_merger.close()            



def iterate_and_categorize(curr_student):

    '''
    Iterates recursivly through all files in student_folder and 
    categorizes them into the correct file list. Returns a list 
    of filled file-lists. Also removes unnecessary symbols in 
    filenames that may screw up sorting of files.

    Parameters
    ----------
    curr_student: Path to a student folder. All files in this folder
    and all subfolders will be categorized.
    '''

    picture_files = []
    pdf_files = []
    doc_files = []
    txt_files = []
    file_lists = (picture_files, pdf_files, doc_files, txt_files)
    

    for dirpath, dirnames, filenames in os.walk(curr_student):

        for filename in filenames:

            filename = remove_unnecassary_symbols(dirpath, filename)

            if not categorise_file(dirpath, filename, file_lists):
                move_file_to_folder(dirpath + "\\" + filename, curr_student)
    
    return file_lists



def merge_files_per_category(curr_student, categorized_files):
    '''
    Checks if files of a format for specific student exist and sorts the 
    list alphabetically. Then it either passses the job to create a 
    merged pdf for all files of a specific format to another method or 
    does it itself, if doing it is short and easy enough. When done, 
    there should be one or many pdf files in the main dir for each student,
    named filename_format_.pdf

    Parameters
    ----------
    curr_student: Path to the folder of the current student.

    categorized_files: List of file-lists that each store the files of
    the correct format
    '''

    picture_files, pdf_files, doc_files, txt_files = categorized_files

    # First sort list and then generate pdf from its elements, either
    # by calling a method or doing it here.
    if len(picture_files) != 0 :

        picture_files.sort()

        pic_to_pdf(curr_student, picture_files)

    if len(pdf_files) != 0 :

        pdf_files.sort()

        pdf_merger = PdfFileMerger(strict=False)
        for pdf in pdf_files:
            pdf_merger.append(pdf)

        pdf_merger.write(curr_student + "_pdfs_" + ".pdf")
        pdf_merger.close()

    if len(doc_files) != 0:

        doc_files.sort()

        doc_to_pdf(curr_student, doc_files)
    



def merge_categories(curr_student):
    '''
    Merge all generated pdf files for different file formats
    into one combined pdf file. Name of that pdf file should 
    be reversed folder name, because student folders are named 
    "firstname lastname" and generated pdf should be named 
    "lastname, fistname"

    Parameters
    ----------
    curr_student: Path to folder of student, whose pdfs will
    be merged. Folder wont be used, only path is necessary
    for naming pdf.
    '''

    student_name_reversed = (
        (os.path.basename(curr_student)).split(" ")[1] + ", " + 
        (os.path.basename(curr_student)).split(" ")[0])


    fullmerger = PdfFileMerger(strict=False)

    at_least_one_file = False
    # add existing merged file format specific pdfs to merger
    if os.path.isfile(curr_student + "_pdfs_" + ".pdf"):
        fullmerger.append(curr_student + "_pdfs_" + ".pdf")
        at_least_one_file = True

    if os.path.isfile(curr_student + "_img_" + ".pdf"):
        fullmerger.append(curr_student + "_img_" + ".pdf")
        at_least_one_file = True

    if os.path.isfile(curr_student + "_docs_" + ".pdf"):
        fullmerger.append(curr_student + "_docs_" + ".pdf")
        at_least_one_file = True

    
    if at_least_one_file:
        # merge pdfs and create pdf with reversed name
        fullmerger.write(os.path.dirname(curr_student) + "\\" + 
            student_name_reversed  + ".pdf")
    fullmerger.close()

    
    # delete merged file format specific pdfs
    if os.path.isfile(curr_student + "_pdfs_" + ".pdf"):
        os.remove(curr_student + "_pdfs_" + ".pdf")

    if os.path.isfile(curr_student + "_img_" + ".pdf"):
        os.remove(curr_student + "_img_" + ".pdf")
    
    if os.path.isfile(curr_student + "_docs_" + ".pdf"):
        os.remove(curr_student + "_docs_" + ".pdf")



def create_merged_pdfs(dir_name):
    '''
    Iterates through all student folders and calls the methods
    to categorize the files of a stundent, create a pdf for each 
    category and then merge the pdf of each category into one 
    combined pdf for that stundent. The combined pdf will not be 
    placed in the student dir, but in the main dir, since stundent
    dirs will be deleted.

    Parameters
    ----------
    dir_name: Main dir, into which the merged pdfs will be placed.
    Must contain subfolders for each student.
    '''

    # List of student dirs
    student_folders = [f.path for f in os.scandir(dir_name) if f.is_dir()]


    for student_folder in student_folders:

        categorized_files = iterate_and_categorize(student_folder)

        merge_files_per_category(student_folder, categorized_files)

        merge_categories(student_folder)

        # Remove student dir (cleanup)
        shutil.rmtree(student_folder)



if __name__ == "__main__":
    '''
    Reistee Extracts IServ - Teachers Easy Extractor
    Autor: Marc Franke
    Version: 0.00.787
    
    With reistee you can easily extract student solution zips
    downloaded from IServ.
    '''

    # Get name of folder or zip passed as command line argument
    main_name_w_ext = os.path.basename(sys.argv[1])
    main_name_wo_ext, main_ext = os.path.splitext(main_name_w_ext)

    # If zip is passed, extract it
    if(sys.argv[1].endswith(".zip")):

        extract_dir_name = main_name_wo_ext + " - reistee"
        with zipfile.ZipFile(sys.argv[1], 'r') as zip_ref:
            zip_ref.extractall(extract_dir_name)


    # If folder is passed, copy contents
    if os.path.isdir(sys.argv[1]):

        extract_dir_name = main_name_wo_ext + " - reistee"
        shutil.copytree(sys.argv[1], extract_dir_name)

    create_merged_pdfs(extract_dir_name)


