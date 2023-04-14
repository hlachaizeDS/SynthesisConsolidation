from excel_handling import *
from get_files_from_explorer import *
from SynthesisPlate import *

EXCEL_TEMPLATE_PATH="template_NT3.xlsm"
EXCEL_OUTPUT_PATH="output_excel_file.xlsm"

INITIAL_PATH=Path("C:\\Users\\HL\\DNA Script\\Thomas YBERT - SYNTHESIS OPERATIONS")
INITIAL_STR="2201,.xlsm"

#Get data from files
synthesis_plates=[]
store_list(tmp_data,[])
files=chose_files_from_explorer(INITIAL_PATH,INITIAL_STR)

file_index=1
for file in files:
    print("PROCESS FILE "+ str(file_index) +"/" + str(len(files)))
    file_index +=1
    try:
        syn_plate=SynthesisPlate()
        print(file)
        syn_plate.fill(file)
        synthesis_plates.append(syn_plate)
    except:
        print("Couldn't process file N°" + str(file_index))

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
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Link")[1] + 1).value = "Link"
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values,"Link")[1]+1).hyperlink=syn_plate.path.replace("\\","/")

    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Resin Batch")[1] + 1).value = syn_plate.resin_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Resin Qty")[1] + 1).value = syn_plate.resin_quantity
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "A batch")[1] + 1).value = syn_plate.A_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "C batch")[1] + 1).value = syn_plate.C_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "G batch")[1] + 1).value = syn_plate.G_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "T batch")[1] + 1).value = syn_plate.T_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Nuc Conc")[1] + 1).value = syn_plate.nucs_concentration
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Enz")[1] + 1).value = syn_plate.enzyme_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Enz Conc")[1] + 1).value = syn_plate.enzyme_concentration
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Deblock")[1] + 1).value = syn_plate.deblock_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "Wash")[1] + 1).value = syn_plate.wash_batch
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "<19")[1] + 1).value = syn_plate.less_than_19
    #global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "20-39")[1] + 1).value = syn_plate.bet_20_and_39
    #global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, ">40")[1] + 1).value = syn_plate.above_39
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "20-59")[1] + 1).value = syn_plate.bet_20_and_59
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, "60-99")[1] + 1).value = syn_plate.bet_60_and_99
    global_sheet.cell(row=globalRow,column=f_index(global_sheet_values, ">100")[1] + 1).value = syn_plate.above_99
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Cleavage type")[1] + 1).value = syn_plate.PSP
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "EndoV batch")[1] + 1).value = syn_plate.Endo_b
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "EndoV Concentration")[1] + 1).value = syn_plate.Endo_c
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "< 100pmol")[1] + 1).value = syn_plate.m100_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "< 200pmol")[1] + 1).value = syn_plate.m200_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "< 600pmol")[1] + 1).value = syn_plate.m600_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "< 5uM")[1] + 1).value = syn_plate.C5_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "< 8uM")[1] + 1).value = syn_plate.C8_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "< 10uM")[1] + 1).value = syn_plate.C10_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Average Quantity")[1] + 1).value = syn_plate.Quantities_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Min Quantity")[1] + 1).value = syn_plate.Min_quantities_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Max Quantity")[1] + 1).value = syn_plate.Max_quantities_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "CV Quantity")[1] + 1).value = syn_plate.CV_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Average Concentration")[1] + 1).value = syn_plate.Concentrations_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Min Concentration")[1] + 1).value = syn_plate.Min_Concentrations_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Max Concentration")[1] + 1).value = syn_plate.Max_Concentrations_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "CV Concentration")[1] + 1).value = syn_plate.CV_Concentrations_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Average Volume")[1] + 1).value = syn_plate.Volume_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Min Volume")[1] + 1).value = syn_plate.Min_Volume_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Max Volume")[1] + 1).value = syn_plate.Max_Volume_D
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "CV Volume")[1] + 1).value = syn_plate.CV_Volume_D
    #global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Volume Ratio")[1] + 1).value = syn_plate.Volume_Ratio
    #global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "Real volume")[1] + 1).value = syn_plate.Real_volume
    global_sheet.cell(row=globalRow, column=f_index(global_sheet_values, "NGS run")[1] + 1).value = syn_plate.NGS

    globalRow+=1

#final lines
global_sheet.cell(row=globalRow,column=1).value = "AVERAGE"
for col in range(2,f_index(global_sheet_values,"Link")[1]):
    global_sheet.cell(row=globalRow,column=col).value = "=AVERAGE(" + xlCell(3,col) + ":" + xlCell(globalRow-1,col) + ")"


global_sheet.cell(row=globalRow+1,column=1).value = "TOTAL"
total_columns=["Sample Nb","Flagged Nb","Flagged ML Nb","< 100pmol","< 200pmol","< 50% purity","Total failed"]
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

#VOLUME HISTOGRAMS
volume_sheet=workbook.create_sheet("Volume")
bins=[0,10,15,20,25,30,35,40,50]
globalRow=3
for syn_plate in synthesis_plates:
    volume_sheet.cell(row=globalRow,column=1).value=syn_plate.title
    volume_sheet.cell(row=globalRow,column=21).value=syn_plate.title
    if syn_plate.Volumes_all!=None:
        add96Plate(volume_sheet,syn_plate.Volumes_all,globalRow+2,2,"Volume")
        addHistogram(volume_sheet,bins,globalRow+2,16,xlCell(globalRow+3,3),xlCell(globalRow+3+7,3+11),"Volume","")

    globalRow+=16

#CONCENTRATIONS HISTOGRAMS
Concentrations_sheet=workbook.create_sheet("Concentrations")
bins=[0,2,2.5,5,7.5,10,12.5,15,20]
globalRow=3
for syn_plate in synthesis_plates:
    Concentrations_sheet.cell(row=globalRow,column=1).value=syn_plate.title
    Concentrations_sheet.cell(row=globalRow,column=21).value=syn_plate.title
    if syn_plate.concentrations!=None:
        add96Plate(Concentrations_sheet,syn_plate.concentrations,globalRow+2,2,"Concentrations")
        addHistogram(Concentrations_sheet,bins,globalRow+2,16,xlCell(globalRow+3,3),xlCell(globalRow+3+7,3+11),"Concentrations","")

    globalRow+=16


#------------------------------------------
# DATA TAB
data_sheet=workbook.create_sheet("Data")
globalRow_data=2

#Headers
data_sheet.cell(row=1,column=1).value="Run"
data_sheet.cell(row=1, column=1).font = Font(bold=True)
data_sheet.cell(row=1,column=2).value="Expected_Well"
data_sheet.cell(row=1, column=2).font = Font(bold=True)
data_sheet.cell(row=1,column=3).value="Sequence"
data_sheet.cell(row=1, column=3).font = Font(bold=True)
data_sheet.cell(row=1, column=4).value = "Quantity"
data_sheet.cell(row=1, column=4).font = Font(bold=True)
data_sheet.cell(row=1, column=5).value = "Concentration"
data_sheet.cell(row=1, column=5).font = Font(bold=True)
data_sheet.cell(row=1, column=6).value = "Purity"
data_sheet.cell(row=1, column=6).font = Font(bold=True)

for syn_plate in synthesis_plates:
    #data_sheet.cell(row=globalRow_data, column=1).value = syn_plate.title

    cell_row = 0
    expected_well=1
    if syn_plate.sequences != None:
        for value in syn_plate.sequences.cells:
            data_sheet.cell(row=globalRow_data + cell_row, column=1).value = syn_plate.title
            data_sheet.cell(row=globalRow_data + cell_row, column=2).value = cell_row+1
            cell_row += 1

    cell_row = 0
    if syn_plate.sequences != None:
        for value in syn_plate.sequences.cells:
            data_sheet.cell(row=globalRow_data + cell_row, column=3).value = value
            cell_row += 1

    cell_row = 0
    if syn_plate.quantities != None:
        for value in syn_plate.quantities.cells:
            data_sheet.cell(row=globalRow_data + cell_row, column=4).value = value
            cell_row += 1

    cell_row = 0
    if syn_plate.concentrations != None:
        for value in syn_plate.concentrations.cells:
            data_sheet.cell(row=globalRow_data + cell_row, column=5).value = value
            cell_row += 1

    cell_row = 0
    if syn_plate.purities != None:
        for value in syn_plate.purities.cells:
            data_sheet.cell(row=globalRow_data + cell_row, column=6).value = value
            cell_row += 1


    globalRow_data += syn_plate.nb_samples

#add96Plate(worksheet,[i+1 for i in range(96)],3,1,'Quantities')
#addHistogram(worksheet,[0,10,20,30,40,50,60,70,80,90],3,15,"B4","M11","TestTitle")
workbook.save(filename=EXCEL_OUTPUT_PATH)