# GEToperant
A General Extraction Tool for Med-PC data. GEToperant will export data from your Med-PC data files. It will produce an Excel file with the data labelled according to any custom profile.

# How to Use GEToperant

Using GEToperant involves four steps.
1. Create a data profile
2. Use the checkboxes to select which headers you wish to export
3. Click on the button that corresponds to the output you want
4. Follow the prompts to select your data profile and files

Your data profile tells GEToperant what data you want extracted and what to label each element as. You can extract:
* A single element
* A section of an array
* A whole array

You can also use MPC2XL Row Profiles (MRPs) to extract your data or convert an MRP to an GEToperant profile.

Your data profile needs to have up to 7 pieces of information:
1. A Label
2. A Label Start Value
3. A Label Increment
4. An Array or Variable
5. The Start Element
6. The Increment Element
7. The Stop Element

In order to extract a single element you will need to define:
* The Label
* The Array or Variable
* The Start Element (i.e. the element you want extracted)
* The Increment Element (which must equal 0)

For example, the element A(0) contains the total lever responses. You would define the label as 'Lever Presses', the Array as 'A', the Start Element as 0 and the Increment Element as 0. This tells GEToperant to get the element A(0) from all sessions in the data files you load and to label it 'Lever Presses'.

In order to extract a section of an array you need:
* The label
* The Array or Variable
* The Start Element
* The Increment Element
* The Stop Element
You can also use:
* The Label Start Value
* The Label Increment

Your Stop Element must be greater than your Start Element and your Increment Element must be greater than 0. This will tell GEToperant to start at a particular part of the array and keep going up by the increments you define until it reaches the Stop Element. So if you wanted every second value of the B array from beginning to element 30, you would set the Start Element to 0, the Incremenet Element to 2 and the Stop Element to 30.

The Label Increment and Label Start Value are optional and allow you to define a value to put at the end of your label. This is useful for a series like timebins. For example, you could have a label of 'Responses Min' with a Label Start Value of 1 and a Label Increment of 1. You would then get 'Responses Min 1', 'Responses Min 2', 'Responses Min 3' and so on.

In order to extract an array until it ends you will need the same details as required to extract a section of an array except you should leave the Stop Element blank or write something in it, such as 'End'. However, any text string will be read as the end of the array.

Session comments are not extracted automatically. In order to extract comments provide:
* The Label
* An Array or Variable with the word 'comment' in it (this is not case sensitive)
* A Start Element and Increment Element of 0

Once you have your data profile, you can select your headers.
All headers are selected by default.

You can export your data as:
1. A single worksheet
2. Separate sheets
3. Separate books

Click on the button corresponding to the type of output you want and GEToperant will display windows to select the appropriate files.

For a single worksheet, GEToperant will save all data to one sheet on one Excel file.

For separate sheets, GEToperant will save each data file in a separate worksheet, but in one Excel file.

For separate books, GEToperant will save each data file in a separate Excel file, named after the file that it corresponds to.

# Feature requests, bugs and issues

GEToperant is pretty stable. No new features are anticipated.

Please feel free to contact me with bugs or raise an issue against this repository.

# Citations are welcome!

Khoo, S. Y. (2021). GEToperant: A General Extraction Tool for Med-PC Data. Figshare. doi: 10.6084/m9.figshare.13697851

# Acknowledgements

GEToperant was made possible thanks to the developers behind Python 3 and the following packages: openpyxl, xlrd, xlsxwriter, pyinstaller and py2app.

# License

GEToperant is free software distributed under an MIT license. You can modify it in any way you like, but there is absolutely no warranty. Please distribute widely to anyone who might find it useful!
