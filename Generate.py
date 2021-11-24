import openpyxl
import qrcode

tableName = "Students.xlsx"


def makeQrCode(code):
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_L,
        box_size=10,
        border=4,
    )
    qr.add_data(str(code))
    qr.make(fit=True)

    img = qr.make_image(fill_color="black", back_color="white")
    return img


wb = openpyxl.load_workbook(tableName, data_only=True)

wb.active = 0

listSheet = wb.active
studentsNum = listSheet['F2'].value
codes = []

for i in range(2, studentsNum + 2):
    codes.append(str(listSheet['A' + str(i)].value).zfill(2) + str(listSheet['C' + str(i)].value) +
                 str(ord(listSheet['D' + str(i)].value)))

for i in range(2, studentsNum + 2):
    print(codes[i - 2])
    makeQrCode(codes[i - 2]).save("Codes/" + str(listSheet['B' + str(i)].value).split()[0] +
                                  "-" + str(listSheet['C' + str(i)].value) +
                                  str(listSheet['D' + str(i)].value) + ".png")
    print("---Done!---")
