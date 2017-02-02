
#import sys
import xlrd
class UR:

    def process(sentence, text_file):
        sentence = sentence.split('-')
        newSentence = []
        j = 0
        for i in range(len(sentence) - 1):
            newSentence.insert(j, sentence[i+1])
            j = j + 1

        lastPosition = len(sentence) - 1
        newSentence.insert(lastPosition, sentence[0])

        j=0
        formatted = ""
        for i in newSentence:
            if j == 0:
                formatted = formatted + i
            else:
                if j == len(newSentence) - 1:
                    formatted = formatted + '.' + i
                else:
                    formatted = formatted + '-' + i

            j = j+ 1

        formatted = formatted + ".URMC-SH.ROCHESTER.EDU"
        text_file.write(formatted + '\n')


def main():
        fileName = input("Enter your input filename : ")
        fileNameOut = input("Enter your output filename : ")
        wb = xlrd.open_workbook(fileName)
        values = []
        for s in wb.sheets():
            for row in range(s.nrows):
                for col in range(s.ncols):
                    value  = (s.cell(row,col).value)
                try:
                    value = str(int(value))
                except ValueError:
                    pass
                finally:
                    values.append(value)

        text_file = open(fileNameOut, "w")
        for i in values:
            UR.process(i, text_file)
        text_file.close()


if __name__ == '__main__':
    main()
