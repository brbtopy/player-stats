from openpyxl import load_workbook
from openpyxl.styles import Alignment
import re
import os
from stat import S_IREAD, S_IRGRP, S_IROTH
from stat import S_IWUSR # Need to add this import to the ones above

workbook = load_workbook(filename="Renaissance FC Player Inventory.xlsx")

filename = "Renaissance FC Player Inventory.xlsx"
os.chmod(filename, S_IWUSR|S_IREAD) # This makes the file read/write for the owner


General_Info_sheet = workbook["General Info"]
Player_Discipline_sheet = workbook["Player Discipline"]
Goal_Involvement_sheet = workbook["Goal Involvement"]

#COMMON ENTRIES IN WORKBOOK
#############################################################################################################
player_name = input("Name of player? ")
player_name = player_name.title()


def counter(directory):
    called = True
    if called:
        count_file = open(directory, "r") # open file in read mode
        count = count_file.read() # read data 
        count_file.close() # close file

        count_file = open(directory, "w") # open file again but in write mode
        count = int(count) + 1 # increase the count value by add 1
        count_file.write(str(count)) # write count to file
        count_file.close() # close file

    return count

counter_file = "counters\\cell_counter.txt"
cell = counter(counter_file)

name_cell = "A" + str(cell)
cell = str(cell)

General_Info_sheet[name_cell] = Player_Discipline_sheet[name_cell] = Goal_Involvement_sheet[name_cell] = player_name

try:
    appearances = int(input(f"How many appearances has {player_name} had? "))
    appearances_cell = "B" + cell
    Player_Discipline_sheet[appearances_cell] = Goal_Involvement_sheet[appearances_cell] = appearances
except ValueError:
    appearances_cell = "B" + cell
    Player_Discipline_sheet[appearances_cell] = Goal_Involvement_sheet[appearances_cell] = 0
    pass

#UPDATING GENERAL INFO WORKSHEET
#############################################################################################################
dob = input(f"What is {player_name}'s date of birth? ")
x = re.split(r"\s", dob)
dob_new = x[0] + " " + x[1].title() + " " + x[2]
dob_cell = "B" + cell
General_Info_sheet[dob_cell] = dob_new

try:
    height = input(f"What is {player_name}'s height in centimeters? ")
    height_cell = "C" + cell
    General_Info_sheet[height_cell] = str(int(height)) + " cm"
    General_Info_sheet[height_cell].alignment = Alignment(horizontal='right', vertical='bottom')
except ValueError:
    pass

player_position = input(f"What is {player_name}'s position(s)? ")
player_position_cell = "D" + cell
General_Info_sheet[player_position_cell] = player_position.upper()

best_position = input(f"What is {player_name}'s best position? ")
best_position_cell = "E" + cell
General_Info_sheet[best_position_cell] = best_position.upper()

favourite_foot = input(f"What is {player_name}'s favourite foot? ")
favourite_foot_cell = "F" + cell
General_Info_sheet[favourite_foot_cell] = favourite_foot.title()

nationality = input(f"What is {player_name}'s nationality? ")
nationality_cell = "G" + cell
General_Info_sheet[nationality_cell] = nationality.title()

birth_place = input(f"What is {player_name}'s birth place? ")
birth_place_cell = "H" + cell
General_Info_sheet[birth_place_cell] = birth_place.title()

contact_addr = input(f"What is {player_name}'s contact address? ")
contact_addr_cell = "I" + cell
General_Info_sheet[contact_addr_cell] = contact_addr
General_Info_sheet[contact_addr_cell].alignment = Alignment(horizontal='right', vertical='bottom')

email = input(f"What is {player_name}'s email address? ")
email_cell = "J" + cell
General_Info_sheet[email_cell] = email.lower()

former_club = input(f"What is {player_name}'s former club? ")
former_club_cell = "K" + cell
General_Info_sheet[former_club_cell] = former_club.title()

ghana_card = input(f"What is {player_name}'s Ghana Card Number? ")
ghana_card_cell = "L" + cell
General_Info_sheet[ghana_card_cell] = ghana_card.upper()

passport_number = input(f"What is {player_name}'s passport number? ")
passport_number_cell = "M" + cell
General_Info_sheet[passport_number_cell] = passport_number.upper()

current_residence = input(f"What is {player_name}'s current place of residence? ")
current_residence_cell = "N" + cell
General_Info_sheet[current_residence_cell] = current_residence.title()

guardian_name = input(f"What is {player_name}'s guardian's name? ")
guardian_name_cell = "O" + cell
General_Info_sheet[guardian_name_cell] = guardian_name.title()

emergency_contact = input(f"What is {player_name}'s emergency contact? ")
emergency_contact_cell = "P" + cell
General_Info_sheet[emergency_contact_cell] = emergency_contact
General_Info_sheet[emergency_contact_cell].alignment = Alignment(horizontal='right', vertical='bottom')

languages = input(f"What languages does {player_name} speak? ")
languages_cell = "Q" + cell
General_Info_sheet[languages_cell] = languages.title()

academic_level = input(f"What is the highest education level {player_name} has achieved? JHS, SHS, or UNI: ")
academic_level_cell = "R" + cell
if academic_level.upper() == "UNI":
    General_Info_sheet[academic_level_cell] = "University"
else:
    General_Info_sheet[academic_level_cell] = academic_level.upper()

academic_institution = input(f"What is the name of the institution {player_name} achieved their highest education? ")
academic_institution_cell = "S" + cell
General_Info_sheet[academic_institution_cell] = academic_institution.title()

print("\n[+] Done with General Info Sheet. Moving on to Player Discipline Sheet ... [+]\n")


#UPDATING PLAYER DISCIPLINE SHEET
#############################################################################################################
try:
    yellow_card = input(f"How many yellow cards does {player_name} have? ")
    yellow_card_cell = "C" + cell
    Player_Discipline_sheet[yellow_card_cell] = int(yellow_card)
except ValueError:
    yellow_card_cell = "C" + cell
    Player_Discipline_sheet[yellow_card_cell] = 0
    pass

try:
    red_card = input(f"How many red cards does {player_name} have? ")
    red_card_cell = "D" + cell
    Player_Discipline_sheet[red_card_cell] = int(red_card)
except ValueError:
    red_card_cell = "D" + cell
    Player_Discipline_sheet[red_card_cell] = 0
    pass

print("\n[+] Done with Player Discipline Sheet. Moving on to Goal Involvement Sheet ... [+]\n")

#UPDATING GOAL INVOLVEMENT SHEET
#############################################################################################################
try:
    goals = input(f"How many goals does {player_name} have? ")
    goals_cell = "C" + cell
    Goal_Involvement_sheet[goals_cell] = int(goals)
except ValueError:
    goals_cell = "C" + cell
    Goal_Involvement_sheet[goals_cell] = 0
    pass

try:
    assists = input(f"How many assists does {player_name} have? ")
    assists_cell = "D" + cell
    Goal_Involvement_sheet[assists_cell] = int(assists)
except ValueError:
    assists_cell = "D" + cell
    Goal_Involvement_sheet[assists_cell] = 0
    pass

try:
    penalties_att = input(f"How many penalties has {player_name} attempted? ")
    penalties_att_cell = "E" + cell
    Goal_Involvement_sheet[penalties_att_cell] = int(penalties_att)
except ValueError:
    penalties_att_cell = "E" + cell
    Goal_Involvement_sheet[penalties_att_cell] = 0
    pass

try:
    penalties_scored = input(f"How many penalties has {player_name} scored? ")
    penalties_scored_cell = "F" + cell
    Goal_Involvement_sheet[penalties_scored_cell] = int(penalties_scored)
except ValueError:
    penalties_scored_cell = "F" + cell
    Goal_Involvement_sheet[penalties_scored_cell] = 0
    pass

print("\n[+] Done with Renaissance Player Inventory Workbook [+]")

workbook.save(filename="Renaissance FC Player Inventory.xlsx")
workbook.close()

os.chmod(filename, S_IREAD|S_IRGRP|S_IROTH)