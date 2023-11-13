import shutil
import csv
import os

# File paths
# Directory of this project
WORKING_DIR = os.path.dirname(os.path.realpath(__file__))
# Path to names file
NAMES_FILE_NAME = WORKING_DIR + "names.csv"
# Path to document xml taken from word template
WORD_TEMPLATE_XML = WORKING_DIR + "document.xml"
# Path to word docx unzipped for subbing edited document.xml into
WORD_TEMPLATE_DIR = WORKING_DIR + "word_template/"
# Path to document.xml in the docx structure
WORD_XML_LOCATION = "word/document.xml"
# Path to results directory, where we put the letters
LETTERS_DIR = WORKING_DIR + "letters/"
# Temp zip name
TEMP_ZIP = "temp"

# String constants
CONSTITUENTS_STRING = "On behalf of your constituents"
CONSTITUENT_STRING = "On behalf of your constituent"


def main():
    lines = []
    senator_student_map = {}

    # Names file has columns First Name, Last Name, Senator By Last Name
    with open(NAMES_FILE_NAME, 'r') as names_file:
        # Create csvreader
        reader = csv.reader(names_file)

        # Extract each row one by one
        for row in reader:
            lines.append(row)

    # Create map of senators to students
    for line in lines:
        names = line
        senator = names[2]
        student_first = names[0]
        student_last = names[1]

        # Check for entry
        if senator in senator_student_map:
            senator_student_map[senator].append(student_first + " " + student_last)
        else:
            senator_student_map[senator] = [student_first + " " + student_last]

    print(senator_student_map.keys())
    print(senator_student_map)

    # Create constituents strings for each senator and replace student list in map
    for senator in senator_student_map:
        constituent_string = ""
        students = senator_student_map[senator]

        # Singular Student
        if len(students) == 1:
            constituent_string = CONSTITUENT_STRING + ", " + students[0]
        # Multiple students
        elif len(students) > 1:
            constituent_string = CONSTITUENTS_STRING

            for i in range(0, len(students)):
                # Final student case
                if i == len(students) - 1:
                    constituent_string += " and " + students[i]
                else:
                    constituent_string += ", " + students[i]

        senator_student_map[senator] = constituent_string

    print(senator_student_map)

    # Get XML from our template and break down by %s
    xml_parts = []
    with open(WORD_TEMPLATE_XML, 'r') as template_file:
        template_string = template_file.read()

        # Find both %s and split on them
        index1 = template_string.index("%s")
        part1 = template_string[:index1]
        xml_parts.append(part1)

        index2 = template_string.index("%s", index1 + 2)
        part2 = template_string[index1+2:index2]
        xml_parts.append(part2)

        part3 = template_string[index2+2:]
        xml_parts.append(part3)

    print(xml_parts)

    # Create word documents from template
    for senator in senator_student_map:
        # Generate xml string
        xml_out = xml_parts[0] + senator + xml_parts[1] + senator_student_map[senator] + xml_parts[2]

        # Overwrite our template document xml file
        with open(WORD_TEMPLATE_DIR + WORD_XML_LOCATION, 'w') as word_xml:
            word_xml.write(xml_out)

        shutil.make_archive(TEMP_ZIP, "zip", root_dir=WORD_TEMPLATE_DIR)
        shutil.move(TEMP_ZIP + ".zip", LETTERS_DIR + "Senator_" + senator + "_Invitation.docx")


if __name__ == "__main__":
    main()
