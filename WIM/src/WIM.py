'''
Created on Nov 5, 2016 <-- Started
WIM - Web Inventory Manager v.0.1
This Application allows to configure every item of the inventory list.
It adds categories, subcategories, brands and phone compatibility fields.
@author: francisco Scovino
'''
import tkinter
from tkinter import ttk, PhotoImage
from tkinter import filedialog
import os.path
import csv
import xml.etree.ElementTree as xmlet
from builtins import str

# Stablish Project and Resources path for images and files
pjct_fldr = lambda x: os.path.abspath(os.path.join(os.path.dirname(__file__), x))
rsc_fldr = pjct_fldr('resources') 

# Form Methods
def browse_raw_file_location():
    '''Open file dialog, select the raw file and add it to the product list'''
    global arr_raw_dict
    # Get the location of the raw file
    var_raw_file_location.set(filedialog.askopenfilename())
    
    try:
        # Open the raw file
        with open(var_raw_file_location.get(), 'r') as rawfile:
            # Convert .csv raw file into a list
            rawlist = csv.reader(rawfile)
            # Append all rows into arr_raw_dict
            for row in rawlist:
                # Jump the first heading row
                if row[0] != 'Brand':
                    arr_raw_dict[row[1]] = {'desc': row[2]}
            
    except Exception as inst:
        print('Error on: browse_raw_file_location():')
        print(inst)
    
def browse_main_file_location():
    '''Open file dialog, select the main file and add it to the product list'''
    global arr_main_dict
    global arr_raw_dict
    global arr_products_dict
    # Get the location of the main file
    var_main_file_location.set(filedialog.askopenfilename())
    # Map all columns
    m_make, m_compatible, m_category, m_subcategory, m_part, m_desc1, m_desc2, m_brand = range(0,8)
    
    try:
        # Open main file
        with open(var_main_file_location.get(), 'r') as mainfile:
            # Convert .csv main file into a list
            mainlist = csv.reader(mainfile)
            # Append all rows into arr_main_dict
            for row in mainlist:
                arr_main_dict[row[m_part]] = {
                        'make': row[m_make],
                        'compatible': row[m_compatible],
                        'category': row[m_category],
                        'subcategory': row[m_subcategory],
                        'desc1': row[m_desc1],
                        'desc2': row[m_desc2],
                        'brand': row[m_brand]
                    }
            
            # Populate arr_products_dict with all items from raw list and   arr_main_dict[part]
            # complete its fields with data from main list
            if len(arr_raw_dict) > 0 and len(arr_main_dict) > 0:
                for part, item in arr_raw_dict.items():
                    if part in arr_main_dict:
                        arr_products_dict[part] = {
                            'make': arr_main_dict[part]['make'],
                            'compatible': arr_main_dict[part]['compatible'],
                            'category': arr_main_dict[part]['category'],
                            'subcategory': arr_main_dict[part]['subcategory'],
                            'part': part,
                            'desc1': arr_main_dict[part]['desc1'],
                            'desc2': arr_main_dict[part]['desc2'],
                            'brand': arr_main_dict[part]['brand']
                            }
                    else:
                        arr_products_dict[part] = {
                            'make': '',
                            'compatible': '',
                            'category': '',
                            'subcategory': '',
                            'part': part,
                            'desc1': item['desc'],
                            'desc2': item['desc'],
                            'brand': ''
                        }
                        
            # Display the data from arr_products_dict onto the products treeview
            update_tree_products()
                    
    except Exception as inst:
        print('Error on: browse_main_file_location():')
        print(inst)
        
def update_tree_make():
    '''Populate the make treeview'''
    global arr_device
    
    # Clear all previous content on treeview
    tree_make.delete(*tree_make.get_children())
    
    value = ''
    for device in arr_device:
        if value != device[0]:
            value = device[0]
            tree_make.insert('', 'end', value, values=[value.replace('_', ' ')])
        
def update_tree_model(event=None):
    '''Display the corresponding models on tree-model according to the make selected'''
    global arr_device
    # Delete current values on tree model
    tree_model.delete(*tree_model.get_children())
    # Get the selected item on Make Treeview
    items_index = tree_make.selection()
    if len(items_index) == 1:
        value = tree_make.item(items_index)['values'][0]
        # Fill the var_make_add field, just for easy to add models
        var_make_add.set(value)
        # Populate the Model Treeview
        for device in arr_device:
            if device[0].replace('_', ' ') == value:
                tree_model.insert('', 'end', device[1], values=[device[1].replace('_', ' ')])
            
def update_tree_category():
    '''Populate the category treeview'''
    global arr_category
    
    # Delete current values on treeview
    tree_category.delete(*tree_category.get_children())
    
    value = ''
    for category in arr_category:
        if value != category[0]:
            value = category[0]
            tree_category.insert('', 'end', value, values=[value.replace('_', ' ')])

def update_tree_subcategory(event=None):
    '''Display the corresponding subcategorirs on treeview according to the category selected'''
    global arr_category
    
    # Delete current values on tree subcategory
    tree_subcategory.delete(*tree_subcategory.get_children())
    # Get the selected item
    items_index = tree_category.selection()
    if len(items_index) == 1:
        value = tree_category.item(items_index)['values'][0]
        # Fill the var_category_add field, just for easy to add subcategory
        var_category_add.set(value)
        # Populate the subcategory treeview
        for category in arr_category:
            if category[0].replace('_', ' ') == value:
                tree_subcategory.insert('', 'end', category[1], values=[category[1].replace('_', ' ')])

def update_tree_brand():
    '''Populate the brand treeview'''
    global arr_brand
    
    # Delete Current values on treeview brand
    tree_brand.delete(*tree_brand.get_children())
    
    for brand in arr_brand:
        tree_brand.insert('', 'end', brand, values=[brand.replace('_', ' ')])
        
def update_tree_products(type_name = 'all'):
    '''Display the products. All, Search or Filter'''
    global arr_products_dict
    # temporary list to hold items already added to the search list
    temp_list = []
    # Delete current values on treeview products
    tree_products.delete(*tree_products.get_children())
    
    color_count = -1
    for part, line in sorted(arr_products_dict.items()):
        color_count += 1
        # Build the item list to insert
        item = [
            line['make'],
            line['compatible'],
            line['category'],
            line['subcategory'],
            part,
            line['desc1'],
            line['brand']
            ]
        
        # Insert all items to the treeview
        if type_name == 'all':
            # Add Even Row Colored in Blue
            if color_count % 2 == 0:
                tree_products.insert('', 'end', values=item, tags = ('even'))
            else:
            # Add Odd Row Colored in White
                tree_products.insert('', 'end', values=item)
                
        # Insert only New Items
        elif type_name == 'filter':
            # Filter ON
            if var_filer_item.get() == 1:
                # Show only items that are missing any value
                if line['make'] == '' or line['compatible'] == '' or line['category'] == '' or line['subcategory'] == '' or line['brand'] == '':
                    # Add Even Row Colored in Blue
                    if color_count % 2 == 0:
                        tree_products.insert('', 'end', values=item, tags = ('even'))
                    else:
                    # Add Odd Row Colored in White
                        tree_products.insert('', 'end', values=item)
                        
            else:
                # Filter OFF
                # Add Even Row Colored in Blue
                if color_count % 2 == 0:
                    tree_products.insert('', 'end', values=item, tags = ('even'))
                else:
                # Add Odd Row Colored in White
                    tree_products.insert('', 'end', values=item)
        
        # Insert only Search Items
        elif type_name == 'search':
            # Make sure the search list has at least one item to search for
            if len(arr_search) > 0:
                checker = 0
                for word in arr_search:
                    if word in str(line['desc1'].lower()) and part not in temp_list:
                        continue
                    else:
                        checker += 1
                # Add item to the treeview
                if checker == 0:
                    # Add part to temp list to keep track
                    temp_list.append(part)
                    # Add Even Row Colored in Blue
                    if color_count % 2 == 0:
                        tree_products.insert('', 'end', values=item, tags = ('even'))
                    else:
                    # Add Odd Row Colored in White
                        tree_products.insert('', 'end', values=item)
        
def move_item_first():
    '''Move the selected item to te first row'''
    global arr_device
    # Get the selected Make
    make_index = tree_make.selection()
    make = tree_make.item(make_index[0])['values'][0]
    # Get the selected item
    model_index = tree_model.selection()
    model = str(tree_model.item(model_index[0])['values'][0])
    # Get the index of the first item on list
    index_first_item = 0
    while arr_device[index_first_item][0] != make or arr_device[index_first_item][2] != 1:
        index_first_item += 1
    # Get the index of the selected item
    index_selected_item = 0    
    for i in range(0, len(arr_device)):
        if arr_device[i][0] == make and arr_device[i][1] == model:
            index_selected_item = i
            
    # Pop out the selected item from its current position    
    item = arr_device.pop(index_selected_item)
    # Increase the position number of items from pos=1 to pos=5
    if item[2] == 5:
        for i in range(0, len(arr_device)):
            if arr_device[i][0] == make and 1 <= arr_device[i][2] <= 4:
                arr_device[i][2] += 1
    elif item[2] == 4:
        for i in range(0, len(arr_device)):
            if arr_device[i][0] == make and 1 <= arr_device[i][2] <= 3:
                arr_device[i][2] += 1
    elif item[2] == 3:
        for i in range(0, len(arr_device)):
            if arr_device[i][0] == make and 1 <= arr_device[i][2] <= 2:
                arr_device[i][2] += 1
    elif item[2] == 2:
        for i in range(0, len(arr_device)):
            if arr_device[i][0] == make and arr_device[i][2] == 1:
                arr_device[i][2] += 1
    # Set the new position for the item
    item[2] = 1
    # Insert the item back to the top position
    arr_device.insert(index_first_item, item)
    # Update the Treeview
    sort_device_list()
    update_tree_model()

def clean_profile():
    ''''Clean selected profile'''
    # Clean Selection on make treeview
    items_index_make = tree_make.selection()
    if len(items_index_make) > 0:
        tree_make.selection_remove(items_index_make)
        
    # Clean Selection on model treeview    
    items_index_model = tree_model.selection()
    if len(items_index_model) > 0:
        tree_model.selection_remove(items_index_model)
    tree_model.delete(*tree_model.get_children())
    
    # Clean Selection on category treeview    
    items_index_category = tree_category.selection()
    if len(items_index_category) > 0:
        tree_category.selection_remove(items_index_category)
        
    # Clean Selection on subcategory treeview    
    items_index_subcategory = tree_subcategory.selection()
    if len(items_index_subcategory) > 0:
        tree_subcategory.selection_remove(items_index_subcategory)
    tree_subcategory.delete(*tree_subcategory.get_children())
    
    # Clean Selection on brand treeview
    items_index_brand = tree_brand.selection()
    if len(items_index_brand) > 0:
        tree_brand.selection_remove(items_index_brand)

def copy_profile():
    '''Pass the selected item's profile to the profile board'''
    # Get the product selected
    index_products = tree_products.selection()
    if len(index_products) == 1:
        item = tree_products.item(index_products)['values']
        # Assign product values to variables
        make = str(item[0]).strip().replace(' ', '_')
        model = str(item[1]).split(' / ')
        category = str(item[2]).strip().replace(' ', '_')
        subcategory = str(item[3]).strip().replace(' ', '_')
        brand = str(item[6]).strip().replace(' ', '_')
        
        try:
            # Select Make
            tree_make.selection_set(make)
            # Select Models
            update_tree_model()
            for m in model:
                tree_model.selection_add([m.strip().replace(' ', '_')])
            # Select Category
            tree_category.selection_set(category)
            # Select Subcategory
            update_tree_subcategory()
            tree_subcategory.selection_set(subcategory)
            # Select Brand
            tree_brand.selection_set(brand)
            
        except Exception as inst:
            print('Error on: copy_profile():')
            print(inst)

def paste_profile():
    '''Copy selected profile from the board to selected items'''
    global arr_products_dict
    
    # Get make value
    make = ''
    index_make = tree_make.selection()
    if len(index_make) == 1:
        make = tree_make.item(index_make[0])['values'][0]
        
    # Get Model Value
    model = []
    index_model = tree_model.selection()
    if len(index_model) > 0:
        for i in range(0, len(index_model)):
            model.append(tree_model.item(index_model[i])['values'][0])
            
    # Get Category value
    category = ''
    index_category = tree_category.selection()
    if len(index_category) == 1:
        category = tree_category.item(index_category[0])['values'][0]
        
    # Get Subcategory value
    subcategory = ''
    index_subcategory = tree_subcategory.selection()
    if len(index_subcategory) == 1:
        subcategory = tree_subcategory.item(index_subcategory[0])['values'][0]
        
    # Get Brand value
    brand = ''
    index_brand = tree_brand.selection()
    if len(index_brand) == 1:
        brand = tree_brand.item(index_brand[0])['values'][0]
        
    # Paste all values to the selected products
    index_products = tree_products.selection()
    if len(index_products) > 0:
        for i in range(0, len(index_products)):
            part = tree_products.item(index_products[i])['values'][4]
            #  Add Values only if there is a selection
            if make != '':
                arr_products_dict[part]['make'] = make.replace('_', ' ')
                
            if len(model) > 0:
                arr_products_dict[part]['compatible'] = ' / '.join(model).replace('_', ' ')
                
            if category != '':
                arr_products_dict[part]['category'] = category.replace('_', ' ')
                
            if subcategory != '':
                arr_products_dict[part]['subcategory'] = subcategory.replace('_', ' ')
                
            if brand != '':
                arr_products_dict[part]['brand'] = brand.replace('_', ' ')
                
        # Update Treeview
        update_tree_products()
        # Release filter
        var_filer_item.set(0)

def insert_or_delete_type(type_name):
    '''Insert or delete an item from the specified type'''
    global arr_device
    global arr_category
    global arr_brand
    
    # Modify Make and Model Treeview
    if type_name == 'model':
        # Add new item to the treeview
        if var_make_add.get() != '' and var_model_add.get() != '':
            line = [str(var_make_add.get()).replace(' ', '_'), str(var_model_add.get()).replace(' ', '_'), 6]
            arr_device.append(line)
            # Clean field
            var_model_add.set('')

        else:
            # Delete selected item from the treeview
            items_index = tree_model.selection()
            if len(items_index) > 0:
                for item in items_index:
                    value = tree_model.item(item)['values'][0]
                    for device in arr_device:
                        if device[0] == var_make_add.get() and str(device[1]) == str(value):
                            line = [var_make_add.get(), str(value), device[2]]
                            arr_device.remove(line)
                            
        # Update Model Treeview
        sort_device_list()
        update_tree_make()
        update_tree_model()
        
    # Modify Category and Subcategory Treeview
    elif type_name == 'subcategory':
        # Add new item to the treeview
        if var_category_add.get() != '' and var_subcategory_add.get() != '':
            line = [var_category_add.get().replace(' ', '_'), var_subcategory_add.get().replace(' ', '_')]
            arr_category.append(line)
            # CLean Field
            var_subcategory_add.set('')
            
        else:
            # Delete selected item from the treeview
            items_index = tree_subcategory.selection()
            if len(items_index) == 1:
                value = tree_subcategory.item(items_index[0])['values'][0]
                line = [var_category_add.get(), value]
                arr_category.remove(line)
                
        # Sort arr_category list
        arr_category.sort(key=lambda x: (x[0], x[1]))
        # Update category and Subcategory treeviews
        update_tree_category()
        update_tree_subcategory()
        
    # Modify Brand Treeview
    elif type_name == 'brand':
        # Add new item to the treeview
        if var_brand_add.get() != '':
            arr_brand.append(var_brand_add.get().replace(' ', '_'))
            # Clean field
            var_brand_add.set('')
        else:
            # Delete seleted item from the treeview
            items_index = tree_brand.selection()
            if len(items_index) == 1:
                value = tree_brand.item(items_index[0])['values'][0]
                arr_brand.remove(value)
                
        # Sort arr_brand list
        arr_brand.sort(key=lambda x: x.lower())
        # Update brand treeview
        update_tree_brand()
        
def sort_device_list():
    '''Sort list keeping the first 5 items of each brand on top'''
    global arr_device
    final_list = []
    
    # Get All Makes on a list
    make_list = []         
    for device in arr_device:
        if device[0] not in make_list:
            make_list.append(device[0])
    
    for make in make_list:
        # Make a temp list for a single brand to process
        temp_list = []
        for device in arr_device:
            if device[0] == make:
                temp_list.append(device)
        
        # Build the top 5 list
        top_list = temp_list[:5]
        # Sort Top List by position
        top_list = sorted(top_list, key=lambda x: x[2])
        # Build the bottom list
        bottom_list = temp_list[5:]
        # Sort bottom list by model
        bottom_list = sorted(bottom_list, key=lambda x: x[1])
        # Add top list to final list
        for item in top_list:
            final_list.append(item)
        # Add Botton list to final list
        for item in bottom_list:
            final_list.append(item)
            
    # Clear arrDevice and add the sorted items
    arr_device[:] = []
    arr_device = final_list
            
def add_keyword_search():
    '''Add keyword and run search filter'''
    global arr_search
    
    # Insert new word to search list
    if var_search.get() != '' and var_search.get() not in arr_search:
        arr_search.append(var_search.get())
        # Change value on clean search button
        btn_clean_search['text'] = 'Clean Filter (' + str(len(arr_search)) + ')'
        # Update Products Treeview
        update_tree_products('search')

def clean_search():
    '''Clean all keywords from search filter'''
    global arr_search
    # Clean the search list
    arr_search[:] = []
    # Clean the field
    var_search.set('')
    # Change value on clean search button
    btn_clean_search['text'] = 'Clean Filter'
    # Update Products Treeview
    update_tree_products()

def export_file():
    '''Save and export product list and profile'''
    global arr_device
    global arr_products_dict
    final_list = []
    added_model_list = []
    added_part_list = []
    
    # Set LOcation to save the exported file
    var_export_path.set(filedialog.asksaveasfilename() + '.csv')
    
    # Get All Makes on a list
    make_list = []         
    for device in arr_device:
        if device[0] not in make_list:
            make_list.append(device[0])
            
    # Make a temp list for a single brand to include in final_list
    for make in make_list:
        model_list = []
        for device in arr_device:
            if device[0] == make:
                model_list.append(device[1])
                
        # Get a list of top5 models
        top_list = model_list[:5]
        # Get the ramaining items from the list
        bottom_list = model_list[5:]
        # Add bottom models first to the final list
        for model in bottom_list:
            for part, line in sorted(arr_products_dict.items()):
                m = model.replace('_', ' ')
                value = str(line['compatible']).split(' / ')
                if len(value) == 1 and m in value and m not in added_model_list and part not in added_part_list:
                    item = []
                    item.append(line['make'])
                    item.append(line['compatible'])
                    item.append(line['category'])
                    item.append(line['subcategory'])
                    item.append(part)
                    item.append(line['desc1'])
                    item.append(line['desc1'])
                    item.append(line['brand'])
                    #Add item to final list
                    final_list.append(item)
                    #Add model to added list for reference
                    added_model_list.append(m)
                    added_part_list.append(part)
                    
        # Add top models to final list            
        for model in top_list:
            for part, line in sorted(arr_products_dict.items()):
                m = model.replace('_', ' ')
                value = str(line['compatible']).split(' / ')
                if len(value) == 1 and m in value and m not in added_model_list and part not in added_part_list:
                    item = []
                    item.append(line['make'])
                    item.append(line['compatible'])
                    item.append(line['category'])
                    item.append(line['subcategory'])
                    item.append(part)
                    item.append(line['desc1'])
                    item.append(line['desc1'])
                    item.append(line['brand'])
                    #Add item to final list
                    final_list.append(item)
                    #Add model to added list for reference
                    added_model_list.append(m)
                    added_part_list.append(part)
                    
    # Add any remaining item to the final list                
    for part, line in sorted(arr_products_dict.items()):
        if part not in added_part_list:
            item = []
            item.append(line['make'])
            item.append(line['compatible'])
            item.append(line['category'])
            item.append(line['subcategory'])
            item.append(part)
            item.append(line['desc1'])
            item.append(line['desc1'])
            item.append(line['brand'])
            #Add item to final list
            final_list.append(item)
            #Add part to added list for reference
            added_part_list.append(part)
        
    # Create a csv writer
    try:
        writer = csv.writer(open(var_export_path.get(), 'w', newline=''))
        for item in final_list:
            writer.writerow(item)
            
    except Exception as inst:
        print('Error on: export_file():')
        print(inst)
    else:
        print('*** List Exported Succesfully')
    
def save_and_close():
    '''Save profile and close application'''
    save_settings()
    root.destroy()

def save_settings():
    '''Save all profile settings to xml file'''
    global arr_device
    global arr_category
    global arr_brand
    
    try:
        data = xmlet.Element('Data')
        device = xmlet.SubElement(data, 'Device')
        category = xmlet.SubElement(data, 'Category')
        brand = xmlet.SubElement(data, 'Brand')
        
        # Create the Devices
        current_make = ''
        for item in arr_device:
            if item[0] != current_make:
                make_name = xmlet.SubElement(device, item[0])
                model_name = xmlet.SubElement(make_name, 'model')
                model_name.text = item[1]
                model_name.attrib['pos'] = str(item[2]) 
                current_make = item[0]
                    
            else:
                 
                for child_of_device in device:
                    if child_of_device.tag == item[0]:
                        model_name = xmlet.SubElement(make_name, 'model')
                        model_name.text = item[1]
                        model_name.attrib['pos'] = str(item[2])
                        
        # Create the Categories
        current_category = ''
        for item in arr_category:
            if item[0] != current_category:
                category_name = xmlet.SubElement(category, item[0])
                subcat_name = xmlet.SubElement(category_name, 'type')
                subcat_name.text = item[1]
                current_category = item[0]
                
            else:
                
                for child_of_category in category:
                    if child_of_category.tag == item[0]:
                        subcat_name = xmlet.SubElement(category_name, 'type')
                        subcat_name.text = item[1]
                        
        # Create the Brands
        for item in arr_brand:
            brand_name = xmlet.SubElement(brand, 'name')
            brand_name.text = item
        
        # Print to file
        tree = xmlet.ElementTree(data)
        tree.write('settings.xml')
        
    except Exception as inst:
        print('Error on: save_settings():')
        print(inst)
        
    else:
        print('.... Settings Succesfully Saved. Good bye...')

def load_settings():
    '''Load profile from xml settings file'''
    global arr_device
    global arr_category
    global arr_brand
    
    try:
        doc = xmlet.parse(rsc_fldr + '/settings.xml')
        data = doc.getroot()
        
        # Populate the Device List, ['make', 'model', 'position']
        node_devices = data.find('Device')
        index = 0
        for make in node_devices:
            make_name = make.tag
            node_model = make.findall('model')
            for model in node_model:
                item = ['','', '']
                item[0] = make_name
                item[1] = str(model.text)
                item[2] = int(model.attrib['pos'])
                # Add item to the list
                arr_device.append(item)
                # Increase the index number
                index += 1
        # Update the make treeviews
        update_tree_make() 
        
        # Populate the Category List, ['category', 'subcategory']
        nodes_category = data.find('Category')
        for genre in nodes_category:
            genre_name = genre.tag#.replace('_', ' ')
            node_value = genre.findall('type')
            for value in node_value:
                item = ['','']
                item[0] = genre_name
                item[1] = value.text
                arr_category.append(item)
        # Update the category treeviews
        update_tree_category()
        
        # Populate the Brand List
        nodes_brand = data.find('Brand')
        for name in nodes_brand:
            arr_brand.append(name.text)
        # Update the brand treeview
        update_tree_brand()
        
    except Exception as inst:
        print('Error on: load_settings():')
        print(inst)
        
# Main Form Layout
root = tkinter.Tk()
root.title('WIM - Web Inventory Manager')
#root.minsize(1368, 768)
frmMain = tkinter.Frame(root).grid(column=0, row=0)

# Form VariablesTreeItem
var_raw_file_location = tkinter.StringVar()
var_main_file_location = tkinter.StringVar()
var_make_add = tkinter.StringVar()
var_model_add = tkinter.StringVar()
var_category_add = tkinter.StringVar()
var_subcategory_add = tkinter.StringVar()
var_brand_add = tkinter.StringVar()
var_search = tkinter.StringVar()
var_filer_item = tkinter.IntVar()
var_export_path = tkinter.StringVar()

arr_raw_dict = {}
arr_main_dict = {}
arr_products_dict = {}
arr_device = []
arr_category = []
arr_brand = []
arr_search = []

# Form Widgets: 7 columns x 11 rows
# Row #0: Raw File
ttk.Label(frmMain, text='Raw File:', width=12, anchor='e').grid(column=0, row=0, padx=5, pady=5, sticky='E')
ttk.Entry(frmMain, textvariable=var_raw_file_location).grid(column=1, columnspan=10, row=0, padx=5, pady=5, sticky='WE')
img_folder_up = PhotoImage(file=rsc_fldr + '/folder_up.png')
ttk.Button(frmMain, text='Browse', image=img_folder_up, compound='left', command=browse_raw_file_location, width=10).grid(column=11, row=0, padx=5, pady=5, sticky='W')

# Row #1: Main File
ttk.Label(frmMain, text='Main File:', width=12, anchor='e').grid(column=0, row=1, padx=5, pady=5, sticky='E')
ttk.Entry(frmMain, textvariable=var_main_file_location).grid(column=1, columnspan=10, row=1, padx=5, pady=5, sticky='WE')
img_folder_dn = PhotoImage(file=rsc_fldr + '/folder_dn.png')
ttk.Button(frmMain, text='Browse', image=img_folder_dn, compound='left', command=browse_main_file_location, width=10).grid(column=11, row=1, padx=5, pady=5, sticky='W')

# Row #2: Separator
ttk.Separator(frmMain, orient=tkinter.constants.HORIZONTAL).grid(column=0, row=2, columnspan=13, padx=5, pady=5, sticky='WE')

# Row #3: Profile Treeviews
tree_make = ttk.Treeview(frmMain)
tree_make.grid(column=1, row=3, rowspan=3, padx=1, pady=5, sticky='WE')
tree_make['columns'] = ('value')
tree_make.column('value', width=120)
tree_make.heading('value', text='Make')
tree_make['show'] = 'headings'
tree_make['height'] = 10
tree_make['selectmode'] = 'browse'
tree_make.bind('<ButtonRelease-1>', update_tree_model)

scroll_tree_make = ttk.Scrollbar(frmMain, orient=tkinter.VERTICAL, command=tree_make.yview)
scroll_tree_make.grid(column=2, row=3, rowspan=3, pady= 5, sticky='WNS')
tree_make['yscroll'] = scroll_tree_make.set

tree_model = ttk.Treeview(frmMain)
tree_model.grid(column=3, row=3, rowspan=3, padx=1, pady=5, sticky='WE')
tree_model['columns'] = ('value')
tree_model.column('value', width=120)
tree_model.heading('value', text='Model', command=move_item_first)
tree_model['show'] = 'headings'
tree_model['height'] = 10

scroll_tree_model = ttk.Scrollbar(frmMain, orient=tkinter.VERTICAL, command=tree_model.yview)
scroll_tree_model.grid(column=4, row=3, rowspan=3, pady= 5, sticky='WNS')
tree_model['yscroll'] = scroll_tree_model.set

tree_category = ttk.Treeview(frmMain)
tree_category.grid(column=5, row=3, rowspan=3, padx=1, pady=5, sticky='WE')
tree_category['columns'] = ('value')
tree_category.column('value', width=120)
tree_category.heading('value', text='Category')
tree_category['show'] = 'headings'
tree_category['height'] = 10
tree_category['selectmode'] = 'browse'
tree_category.bind('<ButtonRelease-1>', update_tree_subcategory)

scroll_tree_category = ttk.Scrollbar(frmMain, orient=tkinter.VERTICAL, command=tree_category.yview)
scroll_tree_category.grid(column=6, row=3, rowspan=3, pady= 5, sticky='WNS')
tree_category['yscroll'] = scroll_tree_category.set

tree_subcategory = ttk.Treeview(frmMain)
tree_subcategory.grid(column=7, row=3, rowspan=3, padx=1, pady=5, sticky='WE')
tree_subcategory['columns'] = ('value')
tree_subcategory.column('value', width=120)
tree_subcategory.heading('value', text='SubCategory')
tree_subcategory['show'] = 'headings'
tree_subcategory['height'] = 10
tree_subcategory['selectmode'] = 'browse'

scroll_tree_subcategory = ttk.Scrollbar(frmMain, orient=tkinter.VERTICAL, command=tree_subcategory.yview)
scroll_tree_subcategory.grid(column=8, row=3, rowspan=3, pady= 5, sticky='WNS')
tree_subcategory['yscroll'] = scroll_tree_subcategory.set

tree_brand = ttk.Treeview(frmMain)
tree_brand.grid(column=9, row=3, rowspan=3, padx=1, pady=5, sticky='WE')
tree_brand['columns'] = ('value')
tree_brand.column('value', width=120)
tree_brand.heading('value', text='Brand')
tree_brand['show'] = 'headings'
tree_brand['height'] = 10
tree_brand['selectmode'] = 'browse'

scroll_tree_brand = ttk.Scrollbar(frmMain, orient=tkinter.VERTICAL, command=tree_brand.yview)
scroll_tree_brand.grid(column=10, row=3, rowspan=3, pady= 5, sticky='WNS')
tree_brand['yscroll'] = scroll_tree_brand.set

# Row #3-7: Profile Buttons
img_bin = PhotoImage(file=rsc_fldr + '/delete.png')
ttk.Button(frmMain, command=clean_profile, text='Clean', image=img_bin, compound='left', width=10).grid(column=11, row=3, padx=5, pady=5, sticky='NW')
img_arrow_up = PhotoImage(file=rsc_fldr + '/list_up.png')
ttk.Button(frmMain, command=copy_profile, text='Copy Up', image=img_arrow_up, compound='left', width=10).grid(column=11, row=6, padx=5, pady=5, sticky='W')
img_arrow_dn = PhotoImage(file=rsc_fldr + '/list_dn.png')
ttk.Button(frmMain, command=paste_profile, text='Copy Dn', image=img_arrow_dn, compound='left', width=10).grid(column=11, row=7, padx=5, pady=5, sticky='W')

img_pencil = PhotoImage(file=rsc_fldr + '/pencil.png')

ttk.Entry(frmMain, textvariable=var_make_add, width=20).grid(column=1, row=6, padx=5, pady=5, sticky='WE')

ttk.Entry(frmMain, textvariable=var_model_add, width=20).grid(column=3, row=6, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Insert | Delete', image=img_pencil, compound='left', command=lambda: insert_or_delete_type('model'), width=10).grid(column=3, row=7, padx=5, pady=5, sticky='WE')

ttk.Entry(frmMain, textvariable=var_category_add, width=20).grid(column=5, row=6, padx=5, pady=5, sticky='WE')

ttk.Entry(frmMain, textvariable=var_subcategory_add, width=20).grid(column=7, row=6, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Insert | Delete', image=img_pencil, compound='left', command=lambda: insert_or_delete_type('subcategory'), width=10).grid(column=7, row=7, padx=5, pady=5, sticky='WE')

ttk.Entry(frmMain, textvariable=var_brand_add, width=20).grid(column=9, row=6, padx=5, pady=5, sticky='WE')
ttk.Button(frmMain, text='Insert | Delete', image=img_pencil, compound='left', command=lambda: insert_or_delete_type('brand'), width=10).grid(column=9, row=7, padx=5, pady=5, sticky='WE')

# Row #8: Separator
ttk.Separator(frmMain, orient=tkinter.constants.HORIZONTAL).grid(column=0, row=8, columnspan=13, padx=5, pady=5, sticky='WE')

# Row #9: Search & Filter Widgets
ttk.Label(frmMain, text='Look for:', width=12, anchor='e').grid(column=0, row=9, padx=5, pady=5, sticky='E')
ttk.Entry(frmMain, textvariable=var_search, width=15).grid(column=1, row=9, padx=5, pady=5, sticky='WE')
img_filter_add = PhotoImage(file=rsc_fldr + '/filter.png')
ttk.Button(frmMain, text='Add Keyword', image=img_filter_add, compound='left', command=add_keyword_search, width=10).grid(column=3, row=9, padx=5, pady=5, sticky='WE')
img_filter_del = PhotoImage(file=rsc_fldr + '/broom.png') 
btn_clean_search = ttk.Button(frmMain, text='Clean Filter', image=img_filter_del, compound='left', command=clean_search, width=10)
btn_clean_search.grid(column=5, row=9, padx=5, pady=5, sticky='WE')

ttk.Checkbutton(frmMain, text='Filter New Items', variable=var_filer_item, command=lambda: update_tree_products('filter'), onvalue=1, offvalue=0).grid(column=11, row=9, padx=5, pady=5, sticky='WE')

# Row #10: Product Treeviiew
tree_products = ttk.Treeview(frmMain)
tree_products.grid(column=0, columnspan=12, row=10, padx=1, pady=1, sticky='WE')
tree_products['columns'] = ('make', 'compatibility', 'category', 'subcategory', 'part', 'description', 'brand')
tree_products.column('make', width=100)
tree_products.column('compatibility', width=200)
tree_products.column('category', width=100)
tree_products.column('subcategory', width=100)
tree_products.column('part', width=150)
tree_products.column('description', width=600)
tree_products.column('brand', width=100)
tree_products.heading('make', text='Make', anchor=tkinter.W)
tree_products.heading('compatibility', text='Compatibility', anchor=tkinter.W)
tree_products.heading('category', text='Category', anchor=tkinter.W)
tree_products.heading('subcategory', text='Subcategory', anchor=tkinter.W)
tree_products.heading('part', text='Part', anchor=tkinter.W)
tree_products.heading('description', text='Description', anchor=tkinter.W)
tree_products.heading('brand', text='Brand', anchor=tkinter.W)
tree_products['show'] = 'headings'
tree_products['height'] = 10
tree_products.tag_configure('even', background='#d5e8ef')

scroll_tree_product = ttk.Scrollbar(frmMain, orient=tkinter.VERTICAL, command=tree_products.yview)
scroll_tree_product.grid(column=12, row=10, sticky='WNS')
tree_products['yscroll'] = scroll_tree_product.set

# Row #11: Products Buttons
img_csv = PhotoImage(file=rsc_fldr + '/csv.png')
ttk.Button(frmMain, text='Export', image=img_csv, compound='left', command=export_file, width=10).grid(column=0, row=11, padx=5, pady=5, sticky='E')
img_exit = PhotoImage(file=rsc_fldr + '/logout.png')
ttk.Button(frmMain, text='Close', image=img_exit, compound='left', command=save_and_close, width=10).grid(column=11, row=11, padx=5, pady=5, sticky='W')

# Load Settings
load_settings()

# Show Form
root.mainloop()
