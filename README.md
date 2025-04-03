<h1 align="center">ADA (Automatic Duty Allocation)</h1>

<p align="center">Automatic duty allocation and attendance sheet generation for school and house prefects at ACS International. Created and maintained by Bono Jakub Gajdek, Secretary/SHOH of the 20th Prefectorial Board.</p>

# Features
- Automatic duty allocations taking into account:
  - Total number of duties
  - Number of duties that week
  - Time since the prefect has last done a specific duty
  - The amount of times a prefect has done a specific duty
- Automatic creation and updating of duty attendance sheets
- Flexible customisation of duties and their requirements based on the house and other circumstances

# Video guide

# Setting up
## Setting up the code
1. Create a new folder for ADA under duties in the Prefectorial Board drive
2. Create a new apps script project (you may call it "ADA")
<img width="400" alt="Screenshot 2025-04-03 at 13 43 53" src="https://github.com/user-attachments/assets/12bac2b6-c948-42d9-afd1-1cc663363756" />

3. Copy the code from [ADA.gs](https://github.com/Bonez07/ADA/blob/main/ADA.gs) and paste it into your apps script project

## Setting up the allocation database
ADA requires all the data of the prefects to be stored in a spreadsheet referred to as the duty allocation database which looks like this:\
<img width="800" alt="Screenshot 2025-04-03 at 22 41 52" src="https://github.com/user-attachments/assets/1ce6acca-b870-443d-b247-7a32665a03f9" />

**1.** Collect the following information about each of the Prefectorial Board members and house prefects who will be doing duty (preferably through a Google Form):
  - Full name
  - Preferred name (the name that will be put on the duty roster)
  - If they are Singaporean (not including PR)
  - If they are Christian
  - Gender
  - Committee (for school prefects only)
  - Whether they are a house prefect or a tutor/year rep (for house prefects only)
> [!NOTE]
> Check with the House Captains of each house what roles there are in the house that are given duties differently. Some houses have no distinction between prefects and year reps, some do, some have more than 2 different roles (e.g. Oldham has dedicated Christian reps to do prayers). Include in the form all options that need to be taken into account when allocating duties.

**2.** Rearrange the data on if they are Singaporean, Christian, their gender and their position into the following 4 letter code format:
- First letter: if they are Singaporean (S for Singaporean, N for non-Singaporean)
- Second letter: if they are Christian (C for christian, N for non-Christian)
- Third letter: their gender (B for boy, G for girl)
- Fourth letter: their position (for school prefects: E for EXCO, S for Subcomm EXCO, H for House Captain, N for other positions) (for house prefects: P for house prefect, Y for year rep)

For example:\
***SCBE*** means the prefect is Singaporean, Christian, a boy and an EXCO\
***NNGN*** means the prefect is non-Singaporean, non-Christian, a girl and a non-EXCO

> [!TIP]
> Use find and replace and the `=CONCATENATE()` function in Google Sheets to easily combine the raw responses to the 4 letter code format

> [!NOTE]
> The specific letters used for the 4 letter codes may be changed and new letters may be used when necessary, as long as the template Google Doc is updated (discussed later on)

**3.** Transfer all the data to a new Google Sheet (you may call it "Duty Allocation Database") with the formatting seen [above](#setting-up-the-allocation-database)

## Setting up attendance sheets


## Setting up roster templates

## Linking the sheets and documents to the code

# Usage instructions



# Developer instructions

# Other notes
What to do when the subcommittee EXCOs are decided?



