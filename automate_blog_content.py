"""
Steps before use the script:
1. Rename .docx file name, to be the correct post title
2. Use pandoc to convert .docx blog post to .md files, copy the .md files to the corresponding topic folder
3. Export images from .docx file to html format, copy the images folder to the same folder as .md files
4. Remove duplicated images in image folders

After using the script:
Copy the images in images folder to target Jekyll image folder

Note:
If header_template_path is not correct, change it as per your need
"""

import re
import os
import os.path
from datetime import datetime
from shutil import copyfile


def rename_file(path):
    """
    This function will convert space in file name to format 'Year-Month-Day-File-Name.md'
    :param path: the path contains .md files
    :return:
    """
    blog_list = os.listdir(path)
    new_file_name = ''
    for blog in blog_list:
        file_name, file_ext = os.path.splitext(blog)
        # Only rename the .md file if there's space in file name
        if 'md' in file_ext and " " in file_name:
            blog_new = blog.replace(" ", "-")
            try:
                content_write_time = read_content_time(os.path.join(path, blog))
                new_file_name = content_write_time + '-' + blog_new
                os.rename(os.path.join(path, blog), os.path.join(path, new_file_name))
                print("Convert file name successfully!")
            except FileExistsError:
                print("Target file with name {0} already exists!".format(new_file_name))
    print("======================================")


def read_content_time(file_path):
    """
    This function will look for the No.3 lines in the .md file and read the time date
    :param file_path: .md file path
    :return: The time date with format 'Year-Month-Day'
    """
    # print(file_path)
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.readlines()
        content_time_str = content[2].rstrip("\n")
        content_time = datetime.strptime(content_time_str, '%A, %B %d, %Y').date()
        return content_time.strftime('%Y-%m-%d')


def generate_header(template_path, file_name):
    """
    This function will generate the jekyll header for the .md file
    :param template_path: The path contains the Jekyll post header template
    :param file_name: The .md file name
    :return: The generated Jekyll header
    """
    file_name_list = file_name.split('-')[3:]
    separator = ' '
    post_title = separator.join(file_name_list)
    with open(template_path, 'r', encoding='utf-8') as f:
        content = f.readlines()
        content[2] = 'title: "{0}"\n'.format(post_title)
        return content


def modify_blog_header(template_path, file_path):
    """
    This function will modify the blog header with the jekyll post header template.
    :param template_path: The path contain the Jekyll post header template
    :param file_path: The .md file path
    :return:
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
        # Return directory without making any change if the file header already matches with Jekyll template
        if '---' in file_lines[0]:
            return
        # Get the first 8 lines of the .md file, No.7 line has 'xa0' means this file contain the author header. \
        # Should be handled differently.
        header_content = file_lines[:8]
        if 'xa0' not in repr(header_content[6]):
            header_content = header_content[:6]
        file_name = f.name.split('\\')[-1].split('.')[0]
        blog_template_header = generate_header(template_path, file_name)
        for i in range(0, len(header_content)):
            if i == len(header_content) - 1:
                file_lines[i] = ''
                for j in range(i, len(blog_template_header)):
                    file_lines[i] += blog_template_header[j]
            else:
                file_lines[i] = blog_template_header[i]
    # Write changes to the .md file
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            for data in file_lines:
                f.write(data)
            f.flush()
    except Exception as e:
        print(e)


def remove_internal_links(file_path):
    """
    This function will remove the internal links start with https://nam06...
    :param file_path: The .md file path
    :return:
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
        # Use regex to match the internal links in line
        pattern = re.compile(r'\(https://.+safelinks\.protection\.outlook\.com.+?\)', re.MULTILINE | re.IGNORECASE)
        match_count = 0
        # Remove the internal links from lines
        for i in range(0, len(file_lines)):
            rest = pattern.findall(file_lines[i])
            if len(rest) != 0:  # match found
                match_count += 1
                for j in range(0, len(rest)):
                    file_lines[i] = file_lines[i].replace(rest[j], '').replace('[', '').replace(']', '')
    # Write changes to the .md file, only do this when match found
    try:
        if match_count > 0:
            with open(file_path, 'w', encoding='utf-8') as f:
                for data in file_lines:
                    f.write(data)
                f.flush()
    except Exception as e:
        print(e)


def remove_start_brackets(file_path):
    """
    This function will remove the '>' at the beginning of a line
    :param file_path: The .md file path
    :return:
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
        pattern = re.compile(r'^>')
        match_count = 0
        # Remove the start angle bracket from lines
        for i in range(0, len(file_lines)):
            rest = pattern.match(file_lines[i])
            if rest is not None:    # match found
                match_count += 1
                file_lines[i] = file_lines[i].replace('>', '', 1)
    # Write changes to the .md file, only do this when match found
    try:
        if match_count > 0:
            with open(file_path, 'w', encoding='utf-8') as f:
                for data in file_lines:
                    f.write(data)
                f.flush()
    except Exception as e:
        print(e)


def get_img_folder_name(file_path):
    """
    This function will return the img folder name of input post by accepting the .md file path
    :param file_path: .md file path
    :return: img folder name
    """
    file_name = os.path.basename(file_path)
    base_path = os.path.dirname(file_path)
    file_name_list = file_name.split('-')[3:]
    separator = ' '
    post_title = separator.join(file_name_list).replace(".md", "")
    img_folder_name = ''
    for item in os.listdir(base_path):
        if os.path.isdir(os.path.join(base_path, item)) and post_title in item.replace("-", " "):
            img_folder_name = item
    return img_folder_name


def get_img_list(file_path):
    """
    This function will return the img list of input post by accepting the .md file path
    :param file_path: .md file path
    :return: img list
    """
    base_path = os.path.dirname(file_path)
    img_folder_name = get_img_folder_name(file_path)
    # print(image_folder_name)
    content_list = os.listdir(os.path.join(base_path, img_folder_name))
    img_list = []
    for item in content_list:
        file_ext = os.path.splitext(item)[1]
        if file_ext in ['.jpg', '.gif', '.png']:
            img_list.append(item)
    return img_list


def rename_img_file(file_path):
    """
    This function will rename the image files from 'image00x' to 'first 20 chars of post name + _image00x'.
    This will ensure image name is unique so that we can add it in to Jekyll blogs.
    This function will also copy the image files into the images folder within path.
    :param file_path: The .md file path
    :return:
    """
    file_name = os.path.basename(file_path)
    base_path = os.path.dirname(file_path)
    img_folder = get_img_folder_name(file_path)
    img_file_list = get_img_list(file_path)
    img_file_list.sort()
    new_file_name = ''
    try:
        for i in range(0, len(img_file_list)):
            new_file_ext = os.path.splitext(img_file_list[i])[1]
            # Don't make any change if the image file name not start with 'image', which means it already has been \
            # renamed.
            if img_file_list[i][:5] != 'image':
                continue
            else:
                # Rename the image file
                new_file_name = img_folder.replace(" ", "_")[0:19] + '_image0' + str(i+1).zfill(2) + new_file_ext
                os.rename(os.path.join(base_path, img_folder, img_file_list[i]), os.path.join(base_path, img_folder,
                                                                                              new_file_name))
                copy_src = os.path.join(base_path, img_folder, new_file_name)
                copy_dst = os.path.join(base_path, 'images', new_file_name)
                # Copy image file to 'images' folder if image does not exist
                if not os.path.exists(copy_dst):
                    copyfile(copy_src, copy_dst)
        print("Rename and copy image files successfully! .md file name: {0}.".format(file_name))
    except FileExistsError:
        print("Target img file with name {0} already exists!".format(new_file_name))
    print("======================================")


def modify_image_link(file_path, target_jekyll_img_folder):
    """
    This function will modify the image links in .md file to format like ![1](/assets/images/img_folder/image.jpg)
    :param file_path: The .md file path
    :param target_jekyll_img_folder: The target img folder in Jekyll site
    :return:
    """
    file_name = os.path.basename(file_path)
    base_path = os.path.dirname(file_path)
    img_list = get_img_list(file_path)
    # Replace image link in .md file with desired format
    with open(file_path, 'r', encoding='utf-8') as f:
        file_lines = f.readlines()
        pattern = re.compile(r'/image', re.IGNORECASE)
        match_count = 0
        matched_line_num = []
        try:
            for i in range(0, len(file_lines)):
                rest = pattern.findall(file_lines[i])
                if len(rest) != 0:  # match found
                    match_count += 1
                    matched_line_num.append(i)
            # Only rename and copy images when match_count = image count
            if match_count != len(img_list):
                print('Error occurred! Img matched count in .md file not equal with img count. Will not perform image '
                      'rename, copy and modify image link in .md file.\nmd file name: {0}\nmatch_count in md: {1}\n'
                      'img count in image folder: {2}'.format(file_name, match_count, len(img_list)))
                print('======================================')
            else:
                rename_img_file(file_path)
                img_list_after_rename = get_img_list(file_path)
                img_list_after_rename.sort()
                for i in range(0, len(matched_line_num)):
                    file_lines[matched_line_num[i]] = "![{0}](/assets/images/{1}/{2})\n"\
                        .format(str(i), target_jekyll_img_folder, img_list_after_rename[i])
        except IndexError as e:
            blog_name = os.path.basename(file_path)
            print("Failed to modify image link for blog: {0}\nLine_No: {1}".format(blog_name, str(i+1)))
            print(e)
    # Write changes to the .md file, only do this when match found
    try:
        if match_count > 0 and (match_count == len(img_list)):
            with open(file_path, 'w', encoding='utf-8') as f:
                for data in file_lines:
                    f.write(data)
                f.flush()
    except Exception as e:
        print(e)


if __name__ == '__main__':
    # Ask for a path contains the .md files and image folders, also prompt to enter a target Jekyll image folder.
    while True:
        blog_path = input("Please enter the path contains the .md files and img folders:")
        jekyll_img_folder_name = input("Please enter the target jekyll site image folder name:")
        if not os.path.exists(blog_path):
            print("The path does not exists, please enter a valid path.")
            os.system('cls')
        else:
            break
    header_template_path = "D:\\AADProjects\\SharedDocs\\Post_header_template.txt"  # This path can be changes as need
    print("Start doing blog content automation...")
    print("======================================")
    print("Start rename blog files...")
    rename_file(blog_path)
    if not os.path.exists(os.path.join(blog_path, 'images')):
        os.makedirs(os.path.join(blog_path, 'images'))
    item_list = os.listdir(blog_path)
    blogs = []
    try:
        print("Start modifying blog header, rename image files, copy images to 'images' folder, remove internal links, "
              "remove start angle brackets, and modify image links...")
        for i in item_list:
            file_ext = os.path.splitext(i)[1]
            if os.path.isfile(os.path.join(blog_path, i)) and "md" in file_ext:
                blogs.append(i)
        for blog_item in blogs:
            blog_item_path = os.path.join(blog_path, blog_item)
            modify_blog_header(header_template_path, blog_item_path)
            remove_internal_links(blog_item_path)
            remove_start_brackets(blog_item_path)
            modify_image_link(blog_item_path, jekyll_img_folder_name)
        print("Successfully completed all work!")
    except Exception as e:
        print(e)




