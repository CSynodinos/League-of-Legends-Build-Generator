import pandas as pd
import xlsxwriter
from tkinter import *
from tkinter import ttk
import tkinter.messagebox

# Function to load Item Spreadsheet.
def Load_Items():
    cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
    rows = 16
    Items = pd.read_excel('Items.xlsx', usecols=cols, nrows=rows)
    Items.head()
    
    return Items

# Function to load Champion spreasheet.
def Load_Champs():
    cols = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]
    rows = 156
    Champ_Base_Stats = pd.read_excel('Champion_Stats.xlsx', usecols=cols, nrows=rows)
    Champ_Base_Stats.head()
    
    return Champ_Base_Stats


### Function that uses the inputs and assigns them as variables.
### By assignin the inputs as globals, they can be used by the Champ_Stats class.
def myInputs():
    global inputChampion
    global lvl
    global inputItem1
    global inputItem2
    global inputItem3
    global inputItem4
    global inputItem5
    global inputItem6
    inputChampion = str(e1.get())
    lvl = int(e0.get())
    inputItem1 = str(e2.get())
    inputItem2 = str(e3.get())
    inputItem3 = str(e4.get())
    inputItem4 = str(e5.get())
    inputItem5 = str(e6.get())
    inputItem6 = str(e7.get())

    ### if/else statement that checks if inputs have been inserted or not.
    if inputChampion:
        root.destroy()
        pass
    else:
        root.destroy()
        tkinter.messagebox.showerror("Error", "Please select a Champion!")

## Pop-up window if the close button (X) is pressed.
def on_closing():
    if tkinter.messagebox.askokcancel("Exit", "Do you want to quit?"):
        root.destroy()
        exit()
    else:
        pass

def base(input, column):
    n = int(column)
    Champ = df2.loc[df2['Champion']==input] # Selected Champion.
    b = Champ.iloc[:,n]
    return b     
def item_select(input_item, column):
    n = int(column)
    Item_Val = df1.loc[df1['Item']==input_item]
    i = int(Item_Val.iloc[:,n])
    return i

def builder():
    # All the variables needed for the calculations in the main builder script.
    # Champion Base Stat Variables
    HP = base(inputChampion, 1)
    HPg = base(inputChampion,2)
    MP = base(inputChampion,3)
    MPg = base(inputChampion,4)
    AD = base(inputChampion,5)
    ADg = base(inputChampion,6)
    AS = base(inputChampion,7)
    ASg = base(inputChampion,8)
    AR = base(inputChampion,9)
    ARg = base(inputChampion,10)
    MR = base(inputChampion,11)
    MRg = base(inputChampion,12)
    MS = base(inputChampion,13)

    # Item Stat Variables. iloc is location based indexing function. int converts location values into integers for calculations.

    ## Item 1. 
    Item_1_AD = item_select(inputItem1, 1)
    Item_1_AP = item_select(inputItem1, 2)
    Item_1_AS = item_select(inputItem1, 3)
    Item_1_HP = item_select(inputItem1, 4)
    Item_1_MP = item_select(inputItem1, 5)
    Item_1_AR = item_select(inputItem1, 6)
    Item_1_MR = item_select(inputItem1, 7)
    Item_1_Critical = item_select(inputItem1, 8)
    Item_1_Life_Steal = item_select(inputItem1, 9)
    Item_1_Haste = item_select(inputItem1, 10)
    Item_1_MS = item_select(inputItem1, 11)

    ## Item 2.
    Item_2_AD = item_select(inputItem2, 1)
    Item_2_AP = item_select(inputItem2, 2)
    Item_2_AS = item_select(inputItem2, 3)
    Item_2_HP = item_select(inputItem2, 4)
    Item_2_MP = item_select(inputItem2, 5)
    Item_2_AR = item_select(inputItem2, 6)
    Item_2_MR = item_select(inputItem2, 7)
    Item_2_Critical = item_select(inputItem2, 8)
    Item_2_Life_Steal = item_select(inputItem2, 9)
    Item_2_Haste = item_select(inputItem2, 10)
    Item_2_MS = item_select(inputItem2, 11)

    ## Item 3.
    Item_3_AD = item_select(inputItem3, 1)
    Item_3_AP = item_select(inputItem3, 2)
    Item_3_AS = item_select(inputItem3, 3)
    Item_3_HP = item_select(inputItem3, 4)
    Item_3_MP = item_select(inputItem3, 5)
    Item_3_AR = item_select(inputItem3, 6)
    Item_3_MR = item_select(inputItem3, 7)
    Item_3_Critical = item_select(inputItem3, 8)
    Item_3_Life_Steal = item_select(inputItem3, 9)
    Item_3_Haste = item_select(inputItem3, 10)
    Item_3_MS = item_select(inputItem3, 11)

    ## Item 4.
    Item_4_AD = item_select(inputItem4, 1)
    Item_4_AP = item_select(inputItem4, 2)
    Item_4_AS = item_select(inputItem4, 3)
    Item_4_HP = item_select(inputItem4, 4)
    Item_4_MP = item_select(inputItem4, 5)
    Item_4_AR = item_select(inputItem4, 6)
    Item_4_MR = item_select(inputItem4, 7)
    Item_4_Critical = item_select(inputItem4, 8)
    Item_4_Life_Steal = item_select(inputItem4, 9)
    Item_4_Haste = item_select(inputItem4, 10)
    Item_4_MS = item_select(inputItem4, 11)

    ## Item 5.
    Item_5_AD = item_select(inputItem5, 1)
    Item_5_AP = item_select(inputItem5, 2)
    Item_5_AS = item_select(inputItem5, 3)
    Item_5_HP = item_select(inputItem5, 4)
    Item_5_MP = item_select(inputItem5, 5)
    Item_5_AR = item_select(inputItem5, 6)
    Item_5_MR = item_select(inputItem5, 7)
    Item_5_Critical = item_select(inputItem5, 8)
    Item_5_Life_Steal = item_select(inputItem5, 9)
    Item_5_Haste = item_select(inputItem5, 10)
    Item_5_MS = item_select(inputItem5, 11)

    ## Item 6.
    Item_6_AD = item_select(inputItem6, 1)
    Item_6_AP = item_select(inputItem6, 2)
    Item_6_AS = item_select(inputItem6, 3)
    Item_6_HP = item_select(inputItem6, 4)
    Item_6_MP = item_select(inputItem6, 5)
    Item_6_AR = item_select(inputItem6, 6)
    Item_6_MR = item_select(inputItem6, 7)
    Item_6_Critical = item_select(inputItem6, 8)
    Item_6_Life_Steal = item_select(inputItem6, 9)
    Item_6_Haste = item_select(inputItem6, 10)
    Item_6_MS = item_select(inputItem6, 11)

    # Base Stat Calculations.

    ## Health.
    Health = HP + HPg*(lvl-1)*(0.7025+0.0175*(lvl-1)) + (Item_1_HP + Item_2_HP + Item_3_HP + Item_4_HP + Item_5_HP + Item_6_HP)

    ## Mana.
    Mana = MP + MPg*(lvl-1)*(0.7025+0.0175*(lvl-1)) + (Item_1_MP + Item_2_MP + Item_3_MP + Item_4_MP + Item_5_MP + Item_6_MP)

    ## Attack Speed.
    Attack_Speed_Base = AS + ASg*(lvl-1)*(0.7025+0.0175*(lvl-1))
    Total_Item_AS = Item_1_AS + Item_2_AS + Item_3_AS + Item_4_AS + Item_5_AS + Item_6_AS

    ### if that checks for Jhin.
    if inputChampion in ('Jhin'):
        Attack_Speed = Attack_Speed_Base
    ### if Attack speed is >2.5, it normalizes at 2.5.
    elif Total_Item_AS >= 210:
        Attack_Speed = 2.5
    elif Total_Item_AS <= 210:
        Attack_Speedt = Attack_Speed_Base * (1 + (Total_Item_AS/100))
        Attack_Speed = float(Attack_Speedt)

    ### Temp Critical and Life Steal Variables.
    Critical_Striket = Item_1_Critical + Item_2_Critical + Item_3_Critical + Item_4_Critical + Item_5_Critical + Item_6_Critical    
    Life_Stealt = Item_1_Life_Steal + Item_2_Life_Steal + Item_3_Life_Steal + Item_4_Life_Steal + Item_5_Life_Steal + Item_6_Life_Steal

    ### Magic Penetration.
    #Magic_Pen = Item_1_Mag_Pen + Item_2_Mag_Pen + Item_3_Mag_Pen + Item_4_Mag_Pen + Item_5_Mag_Pen + Item_6_Mag_Pen

    ### Armor Penetration.

    ### If statement checks if critical and life steal exceed 100%.
    if Critical_Striket > 100 or Life_Stealt > 100:
        Critical_Strike = 100
        Life_Steal = 100
    elif Critical_Striket <= 100 or Life_Stealt <= 100:
        Critical_Strike = Critical_Striket
        Life_Steal = Life_Stealt

    ## Attack Damage.
    ### if checks if the selected champion is Jhin.
    if inputChampion in ('Jhin'):
        ### Function for the counter added inthe ad bonus Jhin gets from lvl ups.      
        def Counter(counter_1):
            counter_1 = counter_1 + lvl
            return counter_1

        ### if sets the bonus lvl ad % according to the input lvl.      
        if 1 <= lvl <= 9: 
            counter_1 = 0
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = 3 + counter_1
        elif 10 <= lvl < 11:
            counter_1 = 0
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = 4 + counter_1
        elif 11 <= lvl < 12:
            counter_1 = 1
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = 4 + counter_1
        elif 12 <= lvl < 13:
            counter_1 = 4
            extra = 4
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        elif 13 <= lvl < 14:
            counter_1 = 4
            extra = 7
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        elif 14 <= lvl < 15:
            counter_1 = 4
            extra = 10
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        elif 15 <= lvl < 16:
            counter_1 = 4
            extra = 13
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        elif 16 <= lvl < 17:
            counter_1 = 4
            extra = 16
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        elif 17 <= lvl < 18:
            counter_1 = 4
            extra = 19
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        elif 18 <= lvl < 19:
            counter_1 = 4
            extra = 22
            for i in range(1):
                counter_1 = Counter(counter_1)
                bonus_lvl_AD = extra + counter_1
        
        ### AD calculations for Jhin.
        Attack_DMG_base = AD + ADg*(lvl-1)*(0.7025+0.0175*(lvl-1))
        Attack_DMG_base_item = Attack_DMG_base + (Item_1_AD + Item_2_AD + Item_3_AD + Item_4_AD + Item_5_AD + Item_6_AD)         
        Attack_DMG_base_item_crit_as = Attack_DMG_base_item + ((Critical_Strike * 0.3) + (Total_Item_AS * 0.25))
        Attack_DMG_lvl = Attack_DMG_base_item_crit_as * (bonus_lvl_AD/100)
        Attack_DMG = Attack_DMG_base_item_crit_as + Attack_DMG_lvl
    else:
        ### AD calculations for all other Champs.
        Attack_DMG = AD + ADg*(lvl-1)*(0.7025+0.0175*(lvl-1)) + (Item_1_AD + Item_2_AD + Item_3_AD + Item_4_AD + Item_5_AD + Item_6_AD)

    ## Ability Power.
    Ability_Power = Item_1_AP + Item_2_AP + Item_3_AP + Item_4_AP + Item_5_AP + Item_6_AP

    ## Champion Armor.
    Armor = AR + ARg*(lvl-1)*(0.7025+0.0175*(lvl-1)) + (Item_1_AR + Item_2_AR + Item_3_AR + Item_4_AR + Item_5_AR + Item_6_AR)

    ## Champion Magic Resistance.
    Magic_R = MR + MRg*(lvl-1)*(0.7025+0.0175*(lvl-1)) + (Item_1_MR + Item_2_MR + Item_3_MR + Item_4_MR + Item_5_MR + Item_6_MR)

    ## Ability Haste.
    Haste = Item_1_Haste + Item_2_Haste + Item_3_Haste + Item_4_Haste + Item_5_Haste + Item_6_Haste

    ## Movement Speed.
    Total_Item_Movement_Speed = Item_1_MS + Item_2_MS + Item_3_MS + Item_4_MS + Item_5_MS + Item_6_MS
    Movement_Speed = MS * (1 + (Total_Item_Movement_Speed/100))

    # Create a workbook and add a worksheet.
    ## Output filename contains Champ name as a string.
    Output_filename = str(inputChampion) + "_Build" + ".xlsx"

    ## Opens an xlsx file to write the output results.
    workbook = xlsxwriter.Workbook(Output_filename)
    worksheet = workbook.add_worksheet()

    ## Structure of output file.
    results = (
        ['Champion', inputChampion],
        ['Level', lvl],
        ['Health', Health],
        ['Mana', Mana],
        ['Attack Damage', Attack_DMG],
        ['Ability Power', Ability_Power],
        ['Attack Speed (%)', Attack_Speed],
        ['Armor', Armor],
        ['Magic Resist', Magic_R],
        ['Critical Strike Chance (%)', Critical_Strike],
        ['Life Steal (%)', Life_Steal],
        ['Ability Haste', Haste],
        ['Movement Speed', Movement_Speed],
        ['Item 1', inputItem1],
        ['Item 2', inputItem2],
        ['Item 3', inputItem3],
        ['Item 4', inputItem4],
        ['Item 5', inputItem5],
        ['Item 6', inputItem6],
    )

    ## Starting rows and columns.
    row = 0
    col = 0

    ## Loop adding the results.
    for Champ, stat in (results):
        worksheet.write(row, col, Champ)
        worksheet.write(row, col + 1, stat)
        row += 1

    ## Close workbook and exit program.
    workbook.close()
    tkinter.messagebox.showinfo("Status", "Build Generated. Press OK to continue")


if __name__== "__main__":
    
    # Indexing of Item Spreadsheet.
    df1 = pd.DataFrame(Load_Items(), columns = ['Item', 'Item_AD', 'Item_AP', 'Item_AS', 'Item_Health', 'Item_Mana', 'Item_AR', 'Item_MR', 
                                                'Item_Critical', 'Item_Life_Steal', 'Item_Haste', 'Item_MS'])

    # Indexing of Champion Spreadsheet.
    df2 = pd.DataFrame(Load_Champs(), columns = ['Champion', 'HP', 'HPg', 'MP', 'MPg', 'AD', 'ADg', 'AS', 'ASg', 'AR', 'ARg', 'MR', 'MRg', 'MS'])

    # The Input pop-up box for Champion, Level and Item selection.
    ## Open an Input box with 8 inputs and label the first input.
    root = Tk()
    root.title("Item Builder")

    ## Label Champion Input.
    labelText1=StringVar()
    labelText1.set("Champion")
    labelDir=Label(root, textvariable=labelText1, height=2)
    labelDir.pack()

    ## Allow for the first input to be written. Input is a string variable.
    directory1=StringVar(None)
    e1=Entry(root,textvariable=directory1,width=50)
    e1.pack()

    ## Label of Level Input.
    labelTextlvl=StringVar()
    labelTextlvl.set("Level")
    labelDir=Label(root, textvariable=labelTextlvl, height=2)
    labelDir.pack()

    ## Allow for the Level input to be written. Input is a string variable.
    directorylvl=StringVar(None)
    e0=Entry(root,textvariable=directorylvl,width=50)
    e0.pack()

    ## Label Item 1 Input.
    labelText2=StringVar()
    labelText2.set("Item 1")
    labelDir=Label(root, textvariable=labelText2, height=2)
    labelDir.pack()

    ## Allow for Item 1 input to be written. Input is a string variable.
    directory2=StringVar(None)
    e2=Entry(root,textvariable=directory2,width=50)
    e2.pack()

    ## Label the Item 2 Input.
    labelText3=StringVar()
    labelText3.set("Item 2")
    labelDir=Label(root, textvariable=labelText3, height=2)
    labelDir.pack()

    ## Allow for Item 2 input to be written. Input is a string variable.
    directory3=StringVar(None)
    e3=Entry(root,textvariable=directory3,width=50)
    e3.pack()


    ## Label the Item 3 Input.
    labelText4=StringVar()
    labelText4.set("Item 3")
    labelDir=Label(root, textvariable=labelText4, height=2)
    labelDir.pack()

    ## Allow for Item 3 input to be written. Input is a string variable.
    directory4=StringVar(None)
    e4=Entry(root,textvariable=directory4,width=50)
    e4.pack()

    ## Label the Item 4 Input.
    labelText5=StringVar()
    labelText5.set("Item 4")
    labelDir=Label(root, textvariable=labelText5, height=2)
    labelDir.pack()

    ## Allow for Item 4 input to be written. Input is a string variable.
    directory5=StringVar(None)
    e5=Entry(root,textvariable=directory5,width=50)
    e5.pack()

    ## Label the Item 5 Input.
    labelText6=StringVar()
    labelText6.set("Item 5")
    labelDir=Label(root, textvariable=labelText6, height=2)
    labelDir.pack()

    ## Allow for Item 5 input to be written. Input is a string variable.
    directory6=StringVar(None)
    e6=Entry(root,textvariable=directory6,width=50)
    e6.pack()

    ## Label the Item 6 Input.
    labelText7=StringVar()
    labelText7.set("Item 6")
    labelDir=Label(root, textvariable=labelText7, height=2)
    labelDir.pack()

    ## Allow for Item 6 input to be written. Input is a string variable.
    directory7=StringVar(None)
    e7=Entry(root,textvariable=directory7,width=50)
    e7.pack()
    
    ## The button to run the code.
    myButton = Button(root, text = "Build", command = lambda: myInputs())
    myButton.pack()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.mainloop()
    builder()