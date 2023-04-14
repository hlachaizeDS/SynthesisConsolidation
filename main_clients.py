from excel_handling import *
from get_files_from_explorer import *
from ClientPlate import *

EXCEL_TEMPLATE_PATH="template_client.xlsm"
EXCEL_OUTPUT_PATH="output_excel_file.xlsm"

INITIAL_PATH=Path("G:\Mon Drive\\0. DNA Script Drive\\3. RD\\RD.2 - DEVELOPMENT\\D.13 - Partner Oligo delivery")
INITIAL_STR="Data file,.xls"

#Get data from files
synthesis_plates=[]
files=chose_files_from_explorer(INITIAL_PATH,INITIAL_STR)

file_index=1
for file in files:
    print("PROCESS FILE "+ str(file_index) +"/" + str(len(files)))
    file_index +=1
    syn_plate=ClientPlate()
    print(file)
    syn_plate.fill(file)
    synthesis_plates.append(syn_plate)

#OPEN WORKBOOK TEMPLATE
workbook=openpyxl.load_workbook(EXCEL_TEMPLATE_PATH,read_only=False,keep_vba=True)
#clear_file(workbook)


#HISTOGRAMS & DATA
histograms_sheet=workbook.create_sheet("Histograms")
data_sheet=workbook.create_sheet("Data")

global_purity_xlrange=""
global_error_rate_xlRange=""
global_quantity_xlrange=""


qty_bins=[0,50,100,150,200,250,300,350,400,450,500]
purity_bins=[0,10,20,30,40,50,60,70,80,90]
error_rate_bins=[0,0.2,0.4,0.6,0.8,1,1.2,1.4,1.6,1.8,2]
globalRow_histograms=3
globalRow_data=2

#Headers
data_sheet.cell(row=1,column=2).value="Quantity"
data_sheet.cell(row=1, column=2).font = Font(bold=True)
data_sheet.cell(row=1, column=3).value = "Purity"
data_sheet.cell(row=1, column=3).font = Font(bold=True)
data_sheet.cell(row=1, column=4).value = "Error Rate"
data_sheet.cell(row=1, column=4).font = Font(bold=True)
data_sheet.cell(row=1, column=5).value = "ML Prediction"
data_sheet.cell(row=1, column=5).font = Font(bold=True)
data_sheet.cell(row=1, column=6).value = "Sequence"
data_sheet.cell(row=1, column=6).font = Font(bold=True)

for client_plate in synthesis_plates:
    #Titles
    data_sheet.cell(row=globalRow_data,column=1).value=client_plate.title

    histograms_sheet.cell(row=globalRow_histograms, column=6).value = client_plate.title
    histograms_sheet.cell(row=globalRow_histograms, column=18).value = client_plate.title
    histograms_sheet.cell(row=globalRow_histograms, column=30).value = client_plate.title
    histograms_sheet.cell(row=globalRow_histograms, column=6).font = Font(bold=True)
    histograms_sheet.cell(row=globalRow_histograms, column=18).font = Font(bold=True)
    histograms_sheet.cell(row=globalRow_histograms, column=30).font = Font(bold=True)

    #Data + histograms
    cell_row = 0
    if client_plate.quantities!=None:

        for value in client_plate.quantities.cells:
            data_sheet.cell(row=globalRow_data+cell_row,column=2).value=value
            cell_row+=1

        addHistogram(histograms_sheet,qty_bins,globalRow_histograms+2,2,xlCell(globalRow_data,2),xlCell(globalRow_data+cell_row,2),"Quantity","Data!")
        client_plate.quantity_xlrange="Data!" + xlCell(globalRow_data,2) + ":" + xlCell(globalRow_data+cell_row,2)
        global_quantity_xlrange+=client_plate.quantity_xlrange + ";"

    cell_row = 0
    if client_plate.purities != None:
        for value in client_plate.purities.cells:
            data_sheet.cell(row=globalRow_data+cell_row,column=3).value = value
            cell_row+=1

        addHistogram(histograms_sheet, purity_bins, globalRow_histograms + 2, 14, xlCell(globalRow_data,3),xlCell(globalRow_data+cell_row,3), "Purity", "Data!")
        client_plate.purity_xlrange = "Data!" + xlCell(globalRow_data,3) + ":" + xlCell(globalRow_data+cell_row,3)
        global_purity_xlrange += client_plate.purity_xlrange + ";"

    cell_row=0
    if client_plate.error_rates != None:
        for value in client_plate.error_rates.cells:
            data_sheet.cell(row=globalRow_data+cell_row,column=4).value = value
            cell_row+=1

        addHistogram(histograms_sheet, error_rate_bins, globalRow_histograms + 2, 26, xlCell(globalRow_data,4),xlCell(globalRow_data+cell_row,4), "Error Rate", "Data!")
        client_plate.error_rate_xlrange = "Data!" + xlCell(globalRow_data,4) + ":" + xlCell(globalRow_data+cell_row,4)
        global_error_rate_xlRange += client_plate.error_rate_xlrange + ";"

    cell_row = 0
    if client_plate.MLflag != None:
        for value in client_plate.MLflag.cells:
            data_sheet.cell(row=globalRow_data+cell_row,column=5).value = value*100
            cell_row += 1

    cell_row = 0
    if client_plate.sequences != None:
        for value in client_plate.sequences.cells:
            data_sheet.cell(row=globalRow_data+cell_row,column=6).value = value
            cell_row += 1




    # histograms_sheet.cell(row=globalRow,column=21).value=syn_plate.title
    # if syn_plate.quantities!=None:
    #     add96Plate(quantity_sheet,syn_plate.quantities,globalRow+2,2,"Quantity")
    #     addHistogram(quantity_sheet,bins,globalRow+2,16,xlCell(globalRow+3,3),xlCell(globalRow+3+8,3+12),"Quantity",strBeforeCells)

    globalRow_histograms+=16
    globalRow_data+=client_plate.nb_samples


#ADD GLOBAL INFO
#worksheet=workbook.create_sheet("Testing",0)
global_sheet=workbook['Global']
global_sheet_values=get_excel_sheet_values(global_sheet)

globalRow=3
for syn_plate in synthesis_plates:
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Client")[1]+1).value=syn_plate.title
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Sample Nb")[1]+1).value=syn_plate.nb_samples
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Max Length")[1]+1).value=syn_plate.max_length
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Flagged Nb")[1]+1).value=syn_plate.failed_flag
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Flagged ML Nb")[1]+1).value=syn_plate.failed_flag_ML
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Link")[1] + 1).value = "Link"
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Link")[1]+1).hyperlink=syn_plate.path.replace("\\","/")
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Link")[1]+1).style="Hyperlink"
    if syn_plate.concentrations!=None:
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Average Concentration")[1] + 1).value = syn_plate.concentrations.mean
    if syn_plate.quantities != None:
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Average Quantity")[1] + 1).value = "=AVERAGE(" + syn_plate.quantity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Min Quantity")[1] + 1).value = "=MIN(" + syn_plate.quantity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Max Quantity")[1] + 1).value = "=MAX(" + syn_plate.quantity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "CV Quantity")[1] + 1).value = "=100*STDEV(" + syn_plate.quantity_xlrange + ")/AVERAGE("+ syn_plate.quantity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "< 200pmol")[1] + 1).value = "=COUNTIF(" + syn_plate.quantity_xlrange + ",\"<200\")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "% failed quantity")[1] + 1).value = "=100*" + xlCell(globalRow,f_index(global_sheet_values,"< 200pmol")[1]+1) + "/" + xlCell(globalRow,f_index(global_sheet_values,"Sample Nb")[1]+1)
    if syn_plate.purities != None:
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Average Purity")[1] + 1).value = "=AVERAGE(" + syn_plate.purity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Average Error rate")[1] + 1).value = "=AVERAGE(" + syn_plate.error_rate_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Min Purity")[1] + 1).value = "=MIN(" + syn_plate.purity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Max Purity")[1] + 1).value = "=MAX(" + syn_plate.purity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "CV Purity")[1] + 1).value = "=100*STDEV(" + syn_plate.purity_xlrange + ")/AVERAGE("+ syn_plate.purity_xlrange + ")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "< 50% purity")[1] + 1).value = "=COUNTIF(" + syn_plate.purity_xlrange + ",\"<50\")"
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "% failed purity")[1] + 1).value = "=100*" + xlCell(globalRow, f_index(global_sheet_values, "< 50% purity")[1] + 1) + "/" + xlCell(globalRow, f_index(global_sheet_values,"Sample Nb")[1] + 1)
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Total % failed")[1] + 1).value = "=" + xlCell(globalRow,f_index(global_sheet_values, "Total % failed")[1] +1-3) + "+" + xlCell(globalRow,f_index(global_sheet_values, "Total % failed")[1] + 1 -1)



    globalRow+=1


#FINAL LINES
global_sheet.cell(row=globalRow,column=1).value = "AVERAGE"
for col in range(2,len(global_sheet_values[0])):
    global_sheet.cell(row=globalRow,column=col).value = "=AVERAGE(" + xlCell(3,col) + ":" + xlCell(globalRow-1,col) + ")"


global_sheet.cell(row=globalRow+1,column=1).value = "TOTAL"
total_columns=["Sample Nb","Flagged Nb","Flagged ML Nb","< 200pmol","< 50% purity"]
for col in range(2,len(global_sheet_values[0])):
    if global_sheet_values[1][col-1] in total_columns:
        global_sheet.cell(row=globalRow+1,column=col).value = "=SUM(" + xlCell(3,col) + ":" + xlCell(globalRow-1,col) + ")"

#add96Plate(worksheet,[i+1 for i in range(96)],3,1,'Quantities')
#addHistogram(worksheet,[0,10,20,30,40,50,60,70,80,90],3,15,"B4","M11","TestTitle")
workbook.save(filename=EXCEL_OUTPUT_PATH)