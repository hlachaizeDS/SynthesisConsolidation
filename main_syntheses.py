from excel_handling import *
from get_files_from_explorer import *
from SynthesisPlate import *

EXCEL_TEMPLATE_PATH="template.xlsm"
EXCEL_OUTPUT_PATH="output_excel_file.xlsm"

INITIAL_PATH=Path("C:\\Users\\HL\\DNA Script\\Thomas YBERT - SYNTHESIS OPERATIONS\\S.3 - P4\\Quartets")
INITIAL_STR="_P4_,.xls"

#Get data from files
synthesis_plates=[]
files=chose_files_from_explorer(INITIAL_PATH,INITIAL_STR)

file_index=1
for file in files:
    print("PROCESS FILE "+ str(file_index) +"/" + str(len(files)))
    file_index +=1
    syn_plate=SynthesisPlate()
    print(file)
    syn_plate.fill(file)
    synthesis_plates.append(syn_plate)

#OPEN WORKBOOK TEMPLATE
workbook=openpyxl.load_workbook(EXCEL_TEMPLATE_PATH,read_only=False,keep_vba=True)
#clear_file(workbook)

#ADD GLOBAL INFO
#worksheet=workbook.create_sheet("Testing",0)
global_sheet=workbook['Global']
global_sheet_values=get_excel_sheet_values(global_sheet)

globalRow=3
for syn_plate in synthesis_plates:
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Plate Name")[1]+1).value=syn_plate.title
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
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Average Quantity")[1] + 1).value = syn_plate.quantities.mean
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Min Quantity")[1] + 1).value = syn_plate.quantities.min
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Max Quantity")[1] + 1).value = syn_plate.quantities.max
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "CV Quantity")[1] + 1).value = syn_plate.quantities.cv
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "< 200pmol")[1] + 1).value = syn_plate.failed_quantity
    if syn_plate.purities != None:
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Average Purity")[1] + 1).value = syn_plate.purities.mean
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Min Purity")[1] + 1).value = syn_plate.purities.min
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Max Purity")[1] + 1).value = syn_plate.purities.max
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "CV Purity")[1] + 1).value = syn_plate.purities.cv
        global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "< 50% purity")[1] + 1).value = syn_plate.failed_purity
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Total failed")[1] + 1).value = syn_plate.failed_purity+syn_plate.failed_quantity

    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Resin Batch")[1] + 1).value = syn_plate.resin_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Resin Qty")[1] + 1).value = syn_plate.resin_quantity
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "A batch")[1] + 1).value = syn_plate.A_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "C batch")[1] + 1).value = syn_plate.C_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "G batch")[1] + 1).value = syn_plate.G_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "T batch")[1] + 1).value = syn_plate.T_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Nuc Conc")[1] + 1).value = syn_plate.nucs_concentration
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Enz")[1] + 1).value = syn_plate.enzyme_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Enz Conc")[1] + 1).value = syn_plate.enzyme_concentration
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Caco")[1] + 1).value = syn_plate.caco_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "HEPES")[1] + 1).value = syn_plate.Hepes_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "CoCl2")[1] + 1).value = syn_plate.coCl2_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "TWEEN-20")[1] + 1).value = syn_plate.t20_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Deblock")[1] + 1).value = syn_plate.deblock_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Wash")[1] + 1).value = syn_plate.wash_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "<19")[1] + 1).value = syn_plate.less_than_19
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "20-39")[1] + 1).value = syn_plate.bet_20_and_39
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, ">40")[1] + 1).value = syn_plate.above_39

    globalRow+=1

#final lines
global_sheet.cell(row=globalRow,column=1).value = "AVERAGE"
for col in range(2,f_index(global_sheet_values,"Link")[1]):
    global_sheet.cell(row=globalRow,column=col).value = "=AVERAGE(" + xlCell(3,col) + ":" + xlCell(globalRow-1,col) + ")"


global_sheet.cell(row=globalRow+1,column=1).value = "TOTAL"
total_columns=["Sample Nb","Flagged Nb","Flagged ML Nb","< 200pmol","< 50% purity","Total failed"]
for col in range(2,len(global_sheet_values[0])):
    if global_sheet_values[1][col-1] in total_columns:
        global_sheet.cell(row=globalRow+1,column=col).value = "=SUM(" + xlCell(3,col) + ":" + xlCell(globalRow-1,col) + ")"

#QUANTITY HISTOGRAMS
quantity_sheet=workbook.create_sheet("Quantity")
bins=[0,50,100,150,200,250,300,350,400,450,500]
globalRow=3
for syn_plate in synthesis_plates:
    quantity_sheet.cell(row=globalRow,column=1).value=syn_plate.title
    quantity_sheet.cell(row=globalRow,column=21).value=syn_plate.title
    if syn_plate.quantities!=None:
        add96Plate(quantity_sheet,syn_plate.quantities,globalRow+2,2,"Quantity")
        addHistogram(quantity_sheet,bins,globalRow+2,16,xlCell(globalRow+3,3),xlCell(globalRow+3+7,3+11),"Quantity","")

    globalRow+=16

#PURITY HISTOGRAMS
purity_sheet=workbook.create_sheet("Purity")
bins=[0,10,20,30,40,50,60,70,80,90]
globalRow=3
for syn_plate in synthesis_plates:
    purity_sheet.cell(row=globalRow,column=1).value=syn_plate.title
    purity_sheet.cell(row=globalRow,column=21).value=syn_plate.title
    if syn_plate.purities!=None:
        add96Plate(purity_sheet,syn_plate.purities,globalRow+2,2,"Purity")
        addHistogram(purity_sheet,bins,globalRow+2,16,xlCell(globalRow+3,3),xlCell(globalRow+3+7,3+11),"Purity","")

    globalRow+=16


#add96Plate(worksheet,[i+1 for i in range(96)],3,1,'Quantities')
#addHistogram(worksheet,[0,10,20,30,40,50,60,70,80,90],3,15,"B4","M11","TestTitle")
workbook.save(filename=EXCEL_OUTPUT_PATH)