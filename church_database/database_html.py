import openpyxl
import emoji
import sys

# Setting up Excel Document
wb = openpyxl.load_workbook('church_database.xlsx')
wb.sheetnames
sheet = wb['Sheet1']

# /Users/jeffreyliew/Documents/GitHub/Church-Database/church_database/database_html.py > /Users/jeffreyliew/Documents/GitHub/Church-Database/church_database/HTML.txt
sys.stout = open('HTML.txt', 'w')

# Setting up HTML
setUp = "\n<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css?family=Spartan\">  <!-- Access to Google Fonts API --> \n<div style=\"overflow-x:auto;\">\n *: College Ministry Available (Check to filter) <input type=\"checkbox\" id=\"myCheck\" onclick=\"myFunction()\">\n<table id=\"myTable\">\n  <tr class=\"header\">\n    <th style=\"width:16%;\" onclick=\"sortTable(0)\">Name</th>\n    <th style=\"width:16%;\" onclick=\"sortTable(1)\">Sunday Worship Time</th>\n    <th style=\"width:16%;\" onclick=\"sortTable(2)\">Address</th>\n    <th style=\"width:16%;\" onclick=\"sortTable(3)\">Distance (miles)</th>\n    <th style=\"width:16%;\" onclick=\"sortTable(4)\">Denomination</th>\n    <th style=\"width:16%;\" onclick=\"sortTable(5)\">Contact Info</th>\n  </tr>"
print(setUp)

# Initalizing the first value
notEmpty = 1
i = 2
emoji = ('\U0001F392')

# Run through each row
while(notEmpty):
    # Read the row and store in appropiate variables
    print()
    name = sheet['A{}'.format(i)].value
    adds = sheet['B{}'.format(i)].value
    dist = sheet['C{}'.format(i)].value
    deno = sheet['D{}'.format(i)].value
    link = sheet['E{}'.format(i)].value
    cMin = sheet['F{}'.format(i)].value
    cont = sheet['H{}'.format(i)].value
    time = sheet['I{}'.format(i)].value




    # Print out HTML code
    print('    <tr>')
    # Check for college ministry
    if(cMin == 'Y'):
        cMin = emoji
        print('        <td><a href={} target="blank">{}</a>{}</td>'.format(link, name,cMin))

    else:
        print('        <td><a href={} target="blank">{}</a></td>'.format(link, name))
    print('        <td>{}</td>'.format(time))
    print('        <td>{}</td>'.format(adds))
    print('        <td>{}</td>'.format(dist))
    print('        <td>{}</td>'.format(deno))
    print('        <td>{}</td>'.format(cont))
    print('    </tr>')

    i = i + 1
    name = sheet['A{}'.format(i)].value
    if not name:
        notEmpty = 0

print()
print("</table>\n</div>\n")

sys.stdout.close()
