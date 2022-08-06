# importing the necessary libraries
import win32com.client
import os
import time

PowerPoint=win32com.client.Dispatch("PowerPoint.Application")
PowerPoint.Visible = True
files_list = []
dir = "C:\\Users\\firstname.lastname\\Folder" #replace this with your path to the pptx documents
file_counter = 1


# start timer
start = time.time()

# create a list of files
for filename in os.listdir(dir):
    files_list.append(str(filename))


items_list = ['item 1', 'item 2', 'item 3', 'item 4']


# iterate over each file, each slide, each shape/table object and search for items in items_list
for file in files_list:
    pptx_filepath = os.path.join(dir, file)
    presentation = PowerPoint.Presentations.Open(pptx_filepath)

    used_items = []
    name = presentation.Name
    name = name.replace('.pptx', '')

    # print which file is currently being searched
    print('-------------------')
    print('{} out of {}; Working on {}...'.format(file_counter, len(files_list), name))

    for item in items_list:
        print('Searching for {}...'.format(str(item)))
        search_text = item
        n = 1
        while n <= presentation.Slides.Count:
            slide = presentation.Slides(n)
            for x in slide.Shapes.Range():

                # handling table objects
                if x.HasTable == -1:
                    # identify table size
                    numofrows = x.Table.Rows.Count
                    numofcols = x.Table.Columns.Count
                    r = 1
                    c = 1

                    # loop over each cell and search for an item
                    while c <= numofcols:
                        r = 1
                        while r <= numofrows:
                            result = x.Table.Cell(r, c).Shape.TextFrame.TextRange.Find(FindWhat=search_text, MatchCase=0)
                            if result is not None:
                                used_items.append(str(result))
                            r += 1
                        c += 1

                # handling shape objects
                if x.HasTextFrame == -1:
                    shape_text = x.TextFrame.TextRange
                    result = shape_text.Find(FindWhat=search_text, MatchCase=0)
                    if result is not None:
                        used_items.append(str(result))
            n += 1

    # close PowerPoint instance
    presentation.Close()

    # update file counter
    file_counter += 1

    # rename the file
    new_name = str(name) + ' done' + '.pptx'
    os.rename(pptx_filepath, os.path.join(dir, new_name))

    # convert each item to lower case
    used_items = list(map(str.lower, used_items))

    # dropping duplicates in items list
    used_items_no_duplicates = list(set(used_items))

    # capitalize each word
    used_items_no_duplicates = list(map(str.title, used_items_no_duplicates))

    print('-------------------')
    print('Found items:')
    print(used_items_no_duplicates)

    #filling out log csv with the file name and found items
    items_str = '; '.join(used_items_no_duplicates)

    with open('log.csv', 'a') as log:
        log.write("{},{}".format(name, items_str))
        log.write("\n")


# calculate and print elapsed time
end = time.time()
elapsed_time = round((end - start) / 60, 2)

print('------------------')
print('elapsed time (min):')
print(elapsed_time)


PowerPoint.Quit()