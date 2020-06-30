import openpyxl
import emoji

# Setting up Excel Document
wb = openpyxl.load_workbook('church_database.xlsx')
wb.sheetnames
sheet = wb['Sheet1']

# Initalizing the first value
name = "Begin"
i = 2
emoji = ('\U0001F392')
# Run through each row
while(len(name) != 0):
    # Read the row and store in appropiate variables
    print()
    name = sheet['A{}'.format(i)].value
    dist = sheet['C{}'.format(i)].value
    deno = sheet['D{}'.format(i)].value
    link = sheet['E{}'.format(i)].value
    cMin = sheet['F{}'.format(i)].value
    cont = sheet['H{}'.format(i)].value
    time = sheet['I{}'.format(i)].value

    # Check for college ministry
    if(cMin == 'Y'):
        cMin = emoji
    elif(cMin == 'N'):
        cMin = ''

    # Print out HTML code
    print('<tr>')
    print('    <td><a href={} target="blank">{}</a>{}</td>'.format(link, name,cMin))
    print('    <td>{}</td>'.format(time))
    print('    <td>{}</td>'.format(dist))
    print('    <td>{}</td>'.format(deno))
    print('    <td>{}</td>'.format(cont))
    print('</tr>')

    i = i +1
    if(i == 4):
        name = ""
