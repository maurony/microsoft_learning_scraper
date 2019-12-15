# import libraries
import argparse
import requests
import os
import pypandoc
import json
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
    
    # create file name
    file_name = course_name.replace('-', '_')

    # specify the url
    module_root = "https://docs.microsoft.com/en-us/learn/"
    root_page = "https://docs.microsoft.com/en-us/learn/paths/" + course_name

    # query the website and return the html to the variable ‘page’
    response = requests.get(root_page)

    # parse the html using beautiful soup and store in variable `soup`
    content = BeautifulSoup(response.text, 'html.parser')

    # get content list
    modules_unit_list = content.find_all('a', {'class':'is-block is-undecorated'})

    # get all sublinks
    modules = dict()

    for index, module in enumerate(modules_unit_list):
        # create new dict
        module_dict = dict()
        # get title and url of current mocule
        module_dict['title'] = module.findChildren('h3')[0].contents[0]
        module_dict['url'] = module_root + module['href'][6:-5]
        # get content of module
        module_response = requests.get(module_dict['url'])
        # get content
        module_content = BeautifulSoup(module_response.text, 'html.parser')
        # module unit list
        module_dict['units'] = [unit['href'] for unit in module_content.find_all('a', {'class':'unit-title'})]
        # add to module
        modules[index] = module_dict
    
    config_file = os.path.join('./config', file_name + '.json')
    if not os.path.exists('./config/'):
        os.makedirs('./config/')
    elif os.path.exists(config_file):
        os.remove(config_file) 
    
    # write config file to json file
    with open(config_file, 'w') as c:
        json.dump(modules, c, indent=1)

    # -----------------------------------------------------------------------
    # get content of entire course
    # create data dir if not exists
    in_file = './data/tmp.html'
    out_file = os.path.join('./data', file_name + '.docx')

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
        for key, module in modules.items():
            print(module['title'] + '-------------------------------')
            f.writelines('<h1>' + module['title'] + '</h1>')
            # iterate over untils
            for unit in module['units']:
                # get url
                unit_url = os.path.join(module['url'] + unit)
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
    pypandoc.convert_file(
        source_file=in_file, 
        format='html', 
        to='docx', \
        outputfile=out_file, 
        extra_args=["+RTS", "-K64m", "-RTS"]
    )

    # remove temp file
    os.remove(in_file)