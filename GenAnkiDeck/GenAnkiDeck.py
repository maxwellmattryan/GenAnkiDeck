# libraries
import random
import genanki
import xlrd

# get deck information and create
def getDeck():
    # get deck ID
    id = random.randrange(1 << 30, 1 << 31)

    # get deck name
    name = input("Please enter a deck name: ").strip()

    # open / write / close the text file opened earlier
    textFile = open("anki_info.txt", "a")
    textFile.write("Deck Name: ".ljust(15) + name + 
                     "\n" + "Deck ID: ".ljust(15) + str(id) +
                     "\n")
    textFile.close()
    
    # create and return deck
    deck = genanki.Deck(id, name)
    return(deck)

# get model information and create
def getModel():
    # generates random model ID
    id = random.randrange(1 << 30, 1 << 31)

    # name input
    name = input("Please enter a model name (note type): ").strip()

    # field and template input and handling
    fieldInput = input("Please enter fields (separated by spaces): ").strip().split(" ")
    fields = [{'name': field} for field in fieldInput]
    afmtString = "{{FrontSide}}<hr id=\"answer\">"
    for field in fieldInput[1:]:
        afmtString += "{{" + field + "}}"
    templates = [
        {
            'name': 'Card 1',
            'qfmt': '{{' + fieldInput[0] + '}}',
            'afmt': afmtString,
        },
    ]
    # call styling info
    styling = getStyling()

    # open / write / close the text file opened earlier
    textFile = open("anki_info.txt", "a")
    textFile.write("Model Name: ".ljust(15) + name +
                     "\n" + "Model ID: ".ljust(15) + str(id) +
                     "\n" + "Fields: ".ljust(15) + str(fields) +
                     "\n" + "Templates: ".ljust(15) + str(templates) +
                     "\n\n")
    print("\n" + "Deck information saved to \"anki_info.txt\"" + "\n")
    textFile.close()

    # create and return model
    model = genanki.Model(id, name, fields, templates, styling)
    return(model)

# asks user for css styling input
def getStyling():
    fontSize = input("Please enter font size (e.g. 24px): ")
    textAlign = input("Please enter text alignment (left, center, or right): ")
    css = ".card { font-size: " + fontSize + ";\n" + "text-align: " + textAlign + " }"
    return(css)

# read data from spreadsheet
def readWriteData(deck, model):
    filename = input("Please enter a filename and extension: ")
    book = xlrd.open_workbook(filename)
    sheet = book.sheet_by_index(0)
    for i in range(sheet.nrows):
        fields = []
        for j in range(sheet.ncols):
            print(sheet.cell_value(i, j))
            fields.append(sheet.cell_value(i, j) + "<br/><br/>")
        note = genanki.Note(model, fields)
        deck.add_note(note)
        print()
    genanki.Package(deck).write_to_file('output.apkg')
    
# main
def main():
    my_deck = getDeck()
    my_model = getModel()
    readWriteData(my_deck, my_model)

main()
