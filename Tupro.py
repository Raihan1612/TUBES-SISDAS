from lib2to3.pytree import convert
import random
import xlsxwriter

def ProgramNode():
    workbook = xlsxwriter.Workbook("Node.xlsx")
    worksheet = workbook.add_worksheet('Node')
    worksheet.write('A1', 'No')
    worksheet.write('B1', 'Label')
    worksheet.write('C1', 'X')
    worksheet.write('D1', 'Y')
    
    for row in range(n):
        no = row+1
        label = no
        convert_label = 'N'+str(label) 
        x = 5* random.randint(1,20)
        y = 5* random.randint(1,20)

        worksheet.write('A'+str(no+1), no)
        worksheet.write('B'+str(no+1), convert_label)
        worksheet.write('C'+str(no+1), x)
        worksheet.write('D'+str(no+1), y)
        
        print(no, convert_label, x, y)

    workbook.close()

def Tetangga():
    workbook = xlsxwriter.Workbook("Tetangga.xlsx")
    worksheet = workbook.add_worksheet('Tetangga')
    
    huruf = 'A'
    for coloumn in range(n):
        no_baris = coloumn+1
        huruf = chr(ord(huruf)+1)
        convert_huruf = huruf + '1'
        worksheet.write(convert_huruf, 'N'+str(no_baris))
        """
        if no_baris == 1:
            huruf = chr(ord(huruf)+1)
            convert_huruf = huruf + '1'
            worksheet.write(convert_huruf, 'N'+str(no_baris))
        else:
        """


    for row in range(n):
        no_row = row+1
        convert_angka = 'N'+str(no_row)
        worksheet.write('A'+str(no_row+1), convert_angka)  

    workbook.close()


if __name__== "__main__":
    
    #Menjalankan Input n
    n = int(input("Masukkan rentang 10..20 : "))
    while (n < 10) or (n > 20):
        print("Anda Salah memasukkan input")
        n = int(input("Masukkan rentang 10..20 : "))
    
    
    ProgramNode()
    Tetangga()