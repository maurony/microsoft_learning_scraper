# import libraries
import argparse
import requests
import os
import pypandoc
from bs4 import BeautifulSoup

# ------------------------------------------------------------------------
# Parse arguments
parser = argparse.ArgumentParser(
    description='Create learning material for microsoft courses!'
)
# get arguments
parser.add_argument("--course", help="The course to be downloaded.")
args = parser.parse_args()

if __name__ == '__main__':
    # course
    course_name = str(args.course)

    # specify the url
    root_page = "https://docs.microsoft.com/en-us/learn/modules" + "/" + course_name

    # query the website and return the html to the variable ‘page’
    response = requests.get(root_page)

    # parse the html using beautiful soup and store in variable `soup`
    content = BeautifulSoup(response.text, 'html.parser')

    # get content list
    course_unit_list = content.find_all('a', {'class':'unit-title'})

    # get all sublinks
    units = []
    for unit in course_unit_list:
        units.append(unit['href'])

    # -----------------------------------------------------------------------
    # get content of entire course
    # create data dir if not exists
    in_file = './data/tmp.html'
    out_file = os.path.join('./data', course_name.replace('-', '_') + '.docx')

    # remove in_file if exists
    if os.path.exists(out_file):
        os.remove(out_file)

    # take care of out_file directory
    if not os.path.exists('./data/'):
        os.makedirs('./data/')
    elif os.path.exists(out_file):
        os.remove(out_file)


    # parse the html using beautiful soup and store in variable `soup`
    with(open(in_file, "a", encoding='utf-8')) as f:
        # iterate over units
        for unit in units:
            # get url
            unit_url = os.path.join(root_page + "/" + unit)
            # logging
            print('Processing unit: ' + unit_url)
            # query the website and return the html to the variable ‘page’
            response = requests.get(unit_url)
            # parse site
            content = BeautifulSoup(response.text, 'html.parser')
            # select relevant elements
            unit_content = content.find_all('div', {'class':'section is-uniform is-relative'})
            # write to file 
            f.writelines(str(unit_content))

    # converting content into word.docx file
    pypandoc.convert(
        source=in_file, 
        format='html', 
        to='docx', \
        outputfile=out_file, 
        extra_args=["+RTS", "-K64m", "-RTS"]
    )

    os.remove(in_file)



