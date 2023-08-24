import re
from openpyxl import Workbook


def main():
    f = open("input.txt", "r")
    lines = [line.strip() for line in f.readlines()]

    # lines = [item.strip() for line in lines for item in re.split(r'(?=[A-FА-Ж]\.)', line) if item]
    lines = [item.strip() for line in lines for item in
             (re.split(r'(?=[A-FА-Ж]\.)', line) if not line.startswith(('SC.', 'MC.')) else [line]) if item]

    lines = [item.strip() for line in lines for item in re.split(r'(\d+\.)\s+', line) if item]
    with open("formatted_input.txt", "w") as fout:
        for line in lines:
            fout.write(line)
            fout.write("\n")

    question_nr = 0
    questions = []
    current_question = {
        "idx": 0,
        "question": "",
        "answers": [],
        "explanation": ""
    }
    in_question = False
    for line in lines:
        # If it's a blank line
        if len(line) == 0:
            in_question = False
            pass

        # Checks if it's a nr. row
        if len(line.split(".")) > 0 and line.split(".")[0].isnumeric() and "." in line:
            question_nr = line.split(".")[0]
            in_question = False
        elif line.startswith("CM.") or line.startswith("CS.") or line.startswith("SC.") or line.startswith("MC.") or line.startswith("СМ.") or line.startswith("CS.") or line.startswith("SC."):
            # now we're parsing a question
            questions.append(current_question)
            in_question = True
            current_question = {"idx": question_nr, "question": line[3:], "answers": [], "explanation": ""}
        elif re.match(r'^[A-FА-Ж]\.', line):
            if in_question:
                current_question["answers"].append(line[2:])
        else:
            if in_question and 0 < len(current_question["answers"]) < 4:
                current_question["answers"][-1] = current_question["answers"][-1] + " " + line
            else:
                current_question["explanation"] = current_question["explanation"] + line
    f.close()
    idx = 1
    while idx < len(questions) :
       print(questions[idx])
       idx += 1
if __name__ == "__main__":
    main()
