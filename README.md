<h1 align="center">ADA (Automatic Duty Allocation)</h1>

<p align="center">Automatic duty allocation and attendance sheet generation for school and house prefects at ACS International. Created and maintained by Bono Jakub Gajdek, Secretary/SHOH of the 20th Prefectorial Board.</p>

# Overview
ADA's functionality can be summarised in the following diagram:
<p align="center">
<img width="600" alt="ADA overview" src="https://github.com/user-attachments/assets/bfe90b1a-9d8a-4c5d-99dd-98df3c848e48" />
</p>

**ADA works with 4 primary files when allocating duties:**
- [Duty allocation database sheet](#setting-up-the-allocation-database)
  - Sheet containing names, data (if they are Singaporean, Christian, gender and position), past duties, days where prefect is unavailable
  - This data is used when allocating duties
- [2 attendance sheets for PB and House](#setting-up-attendance-sheets)
  - Duty attendance speadsheets in which ADA automatically creates attendance worksheets (Spreadsheet refers to the whole file, worksheet refers to the tab under the spreadsheet)
- [Duty roster template doc](#setting-up-roster-templates)
  - Document which contains all of the duties ADA needs to allocate (these are written in a special code explained later)
  - These duties can be flexibly changed as necessary
  - ADA copies this document to create the duty roster and replaces the duty codes with names it allocated
 
**Duty allocations take into account:**
- Requirements of the duty
- Total number of duties
- Number of duties that week
- Time since the prefect has last done a specific duty
- The amount of times a prefect has done a specific duty

# Video guide

# Setting up
## Setting up the code
1. Create a new folder for ADA under duties in the Prefectorial Board drive
2. Create a new apps script project (you may call it "ADA")
<img width="400" alt="Creating apps script project" src="https://github.com/user-attachments/assets/12bac2b6-c948-42d9-afd1-1cc663363756" />

3. Copy the code from [ADA.gs](https://github.com/Bonez07/ADA/blob/main/ADA.gs) and paste it into your apps script project

## Setting up the allocation database
ADA requires all the data of the prefects to be stored in a spreadsheet referred to as the duty allocation database which looks like this:\
<img width="800" alt="Duty Allocation Database" src="https://github.com/user-attachments/assets/1ce6acca-b870-443d-b247-7a32665a03f9" />

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
<a name="4-letter-codes"></a>
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

**4.** Create a new worksheet in the database spreadsheet for every house and fill it in in the same manner. **The worksheet titles must use these precise names**:

<img width="667" alt="Screenshot 2025-04-07 at 22 18 13" src="https://github.com/user-attachments/assets/240c06b9-49df-4682-842e-5d353557ec95" />


## Setting up attendance sheets
2 attendace spreadsheets will be needed for ADA - one for school prefects and one for house prefects (or probation nominees). The attendance worksheets will be generated from a template when ADA is run.

**1.** On the PB attendance Google Sheet, make a worksheet named "Template"

**2.** Fill in the template worksheet with the following format (probation nominees sheet would also follow this format)

<img width="800" alt="Screenshot 2025-04-07 at 08 38 42" src="https://github.com/user-attachments/assets/108ec224-bb1a-46cb-af45-0bddcb32216b" />

> [!IMPORTANT]
> Names on the template must be in the same rows and same order as in the database. Columns up to H (T. Duties all time) must be in the same format as in the image provided.

**3.** On the House Prefects attendance Google Sheet, make a worksheet named "Template"

**4.** Since the House Prefects change every week, the template should not include any names - they will be filled in by the program using the database:

<img width="800" alt="Screenshot 2025-04-07 at 09 03 55" src="https://github.com/user-attachments/assets/17e596fa-e04d-418b-9488-3651d5663a19" />

## Setting up roster templates
For ADA to know what duties to allocate, duties need to but put down on the roster template using a code in curly brackets for every duty. Every single duty that needs to be allocated needs to be written into the document in this manner.

<img width="500" alt="Screenshot 2025-04-07 at 23 13 26" src="https://github.com/user-attachments/assets/36bd72b9-1325-4fbd-a79c-49dee38591b2" />

Duty codes are split into 3 parts separated by underscores:

<img width="560" alt="Screenshot 2025-04-08 at 00 15 49" src="https://github.com/user-attachments/assets/711c709f-4e94-48f8-b792-9406e4128fd5" />

**First part - 2 letter code indicating the type of duty**\
The codes used for each duty can be seen in the image above. These codes may be changed as long as you are consistent across weeks and if you do change them completely the past duties column in the duty allocation databse must be updated.

**Second part - day of the week**\
1 means Monday, 2 means Tuesday and so on

**Third part - duty requirements**\
Requirements of the duty - in [4 letter code format](#4-letter-codes). Each letter corresponds to a letter in the prefects' 4 letter data. A "*" indicates there is no requirement for that letter and can be anything. A letter indicates the duty requires a person with that particular letter in that slot. When there are multiple options for a certain letter, they should both be put in brackets.

For example:
**GN means:
- It doesn't matter if prefect is Singaporean
- It doesn't matter if prefect is Christian
- Prefect must be a girl
- Prefect must be a non-exco (and not subcomm exco)

SC*(NS) means:
- Prefect must be Singaporean
- Prefect must be Christian
- Gender doens't matter
- Prefect must be either a non-exco or a subcomm exco (i.e. all non-excos)

School prefect duties and house prefect duties are differentiated by colour. Duties coloured in black will be given to school prefects, duties coloured in red will be given to house prefects.

### Paired duties
Sometimes, there are pairs of duties where at least one singaporean/any other requirement is needed. For these cases, the required letter can be put in square brackets (e.g. [S]\*\*N and [S]\*\*(ES)) and the program will randomly select one of the duties to have that requirement while it will be replaced with a "\*" for the rest. If you wish to replace with something other than a "\*", you may use a "/" to put another requirement which you wish to replace the rest of the duties with (e.g. S**[P/Y] and \*C\*[P/Y] means one prefect one year rep) For duties to be paired, they must have square brackets on the same letter and must be on the same day. If there are multiple duty pairs on the same day with the same requirement, the requirement can be tagged with an underscore followed by any letter to differentiate the pairs and the letter after the underscore with be ignored (e.g. S**[P_a], \*C\*[P_a] and **G[P_b], ***[P_b] will be treated as 2 separate duty pairs where one from each pair will be randomly given P as the last letter)

I recommend having multiple templates set-up - one for each house to meet their specific requirements and making new templates whenever there are changes to duties on certain weeks.

### Forcing duties
Sometimes, you may want to manually allocate a prefect instead of letting ADA decide. To force someone to get a certain duty, simply put their preferred name (same one as in database) in brackets in place of the duty requirements. This will put the person on the duty and update the attendance sheet and database sheet.

e.g. {{EX_1_(Bono)}} allocates a Monday EXCO duty to Bono

## Linking the sheets and documents to the code
You may have two separate sets of sheets: one set for testing and one actual set. Whether the actual or test set will be used can be toggled via the "mode" variable:

<img width="339" alt="Screenshot 2025-04-09 at 22 19 10" src="https://github.com/user-attachments/assets/272c6b6d-7bfc-48be-87c8-1e24cb58e8f7" />

The sheets and docs used can be linked by copying the ID into the code.

A file's ID can be found in its link as follows:
<img width="718" alt="Screenshot 2025-04-09 at 10 46 24" src="https://github.com/user-attachments/assets/d55b66a4-29d2-4807-8cf6-93c0a723fc3a" />

And pasted in its corresponding variable in the code:
<img width="919" alt="Screenshot 2025-04-09 at 10 47 24" src="https://github.com/user-attachments/assets/27a361a2-54ae-4f4d-b945-8031e18d102f" />

The rosterFolderID variable refers to a google drive folder where ADA will create a copy of the roster.

# Usage instructions
Once all the set-up is finished and a template as well as all the sheets are linked to the code, running the code is relatively straightforward

<img width="1342" alt="Screenshot 2025-04-09 at 23 40 58" src="https://github.com/user-attachments/assets/e819ed63-b631-469b-b9ce-1bb96cbb6b5f" />

After the code is run, the execution log should show up and inform you if the allocation was successful as well as any errors or warnings.



### Config options

# Developer instructions

# Other notes
What to do when the subcommittee EXCOs are decided?

How to set up probation duty weeks?




