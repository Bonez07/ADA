<h1 align="center">ADA (Automatic Duty Allocation)</h1>

<p align="center">Automatic duty allocation and attendance sheet generation for school and house prefects at ACS International. Created and maintained by Bono Jakub Gajdek, Secretary/SHOH of the 20th Prefectorial Board.</p>

# Video guide
[![Video guide](https://img.youtube.com/vi/AvcMz5czzPs/0.jpg)](https://youtu.be/AvcMz5czzPs)


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
  - Year group (for house prefects only)
  - Whether they are a house prefect or a tutor/year rep (for house prefects only)
> [!NOTE]
> Check with the House Captains of each house what roles there are in the house that are given duties differently. Some houses have no distinction between prefects and year reps, some do, some have more than 2 different roles (e.g. Oldham has dedicated Christian reps to do prayers). Include in the form all options that need to be taken into account when allocating duties.

**2.** Rearrange the data on if they are Singaporean, Christian, their gender and their position into the following 4 letter code format:
<a name="4-letter-codes"></a>
- First letter: if they are Singaporean (S for Singaporean, N for non-Singaporean)
- Second letter: if they are Christian (C for christian, N for non-Christian)
- Third letter: their gender (B for boy, G for girl)
- Fourth letter: their position (for school prefects: E for EXCO, S for Subcomm EXCO, H for House Captain, N for other positions) (for house prefects: P for house prefect, Y for year rep)
- For House Captains only: after the 4 letter code add dash and the 3 letter house name

For example:\
***SCBE*** means the prefect is Singaporean, Christian, a boy and an EXCO\
***NNGN*** means the prefect is non-Singaporean, non-Christian, a girl and a non-EXCO
**SNGH-CKS*** means this is the House Captain of CKS and is Singaporean, non-Christian and a girl

> [!TIP]
> Use find and replace and the `=CONCATENATE()` function in Google Sheets to easily combine the raw responses to the 4 letter code format

> [!NOTE]
> The specific letters used for the 4 letter codes may be changed and new letters may be used when necessary, as long as the template Google Doc is updated (discussed later on)

**3.** Transfer all the data to a new Google Sheet (you may call it "Duty Allocation Database") with the formatting seen [above](#setting-up-the-allocation-database)

**4.** Create a new worksheet in the database spreadsheet for every house and fill it in in the same manner. **The worksheet titles must use these precise names**:

<img width="667" alt="Screenshot 2025-04-07 at 22 18 13" src="https://github.com/user-attachments/assets/240c06b9-49df-4682-842e-5d353557ec95" />

### Unavailable days

If you know in advance certain prefects are unable to do duty on certain days (e.g. due to having something on Tuesday whitespace/committee having a major event on that day), you can tell ADA not to give those people duty by putting down the days under the "Unavailable days" column in the database. 1 corresponds to Monday, 2 to Tuesday and so on.

For example if HBHG has peer leader meetings and can't do duty on Tuesday and Affairs has a major event and requested not to do duty on that week you may put:

<img width="784" alt="Screenshot 2025-04-19 at 12 34 28" src="https://github.com/user-attachments/assets/ffa2a456-9186-4ab8-9ed2-33d06d94b82d" />

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
2 letter codes are used to represent the type of duty that is needed. These codes are for storing past duties in the duty allocation database as well as for the code's internal reference. Most of these codes may be changed but I recommend the following convention:

| **Duty** | **Code** |
| ------- | ------- |
| EXCO on supervision | EX |
| Foyer Duty | FY |
| Foyer Attire Escorts | FA |
| PA System (PA, pledge, prayer) | PA |
| Back Gate | BG |
| Flag Raising | FL |
| Assembly/Chapel Stage | MC |
| Assembly/Chapel Escorts | ES |
| Assembly/Chapel Doors | DO |
| Chapel Attire Escorts | CA |

**Second part - day of the week**\
1 means Monday, 2 means Tuesday and so on

**Third part - duty requirements**\
Requirements of the duty - in [4 letter code format](#4-letter-codes). Each letter corresponds to a letter in the prefects' 4 letter data. A "*" indicates there is no requirement for that letter and can be anything. A letter indicates the duty requires a person with that particular letter in that slot. When there are multiple options for a particular letter, the different options should both be put in brackets.

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
Once all the set-up is finished and a template as well as all the sheets are linked to the code, running the code is relatively straightforward. Do make sure your attendance sheet and database are all up to date before you run the code.

<img width="1342" alt="Screenshot 2025-04-09 at 23 40 58" src="https://github.com/user-attachments/assets/e819ed63-b631-469b-b9ce-1bb96cbb6b5f" />

After the code is run, the execution log should show up and inform you if the allocation was successful as well as any errors or warnings.

I recommend running the code in testing mode in the same conditions before making the real roster just to make sure there are no issues.

## Config options
ADA offers a number of settings when giving duties which allow fine tuning of the allocations:
<img width="1084" alt="Screenshot 2025-04-19 at 15 04 25" src="https://github.com/user-attachments/assets/120a5249-b679-4c5c-afd6-027890e39b28" />

Here is an explanation of each:

**copyRosterTemplate** - when true: ADA will copy the roster template doc into the roster folder and fill in the names on the copy. when false: ADA will fill in the names directly on the roster template doc

**schoolPrefectColor** - color of the school prefect names on the roster template (in CSS color format - in hex and in RGB) - this color will be used to pick out which duties are to be given to school prefects (default value is "#000000" which is black)

**housePrefectColor** - color of the house prefect names on the roster template (in CSS color format - in hex and in RGB) - this color will be used to pick out which duties are to be given to house prefects (default value is "#ff0000" which is red)

**subcommProbabilityScalingFactor** - a value between 0 and 1 dictating how likely subcomm excos are to get exco duties. 0 means they will get no exco duties, 1 means they will get the same amount of exco duties as regualar excos

**yearRepScoreOffset** - a value determining how much less duties year reps (marked with last letter "Y") get compared to other house prefects. 0 means same amount, the more positive the less duties (value of ~5 to make every prefect get 2 duties while year reps get 1 or 2, ~2.1 to make all year reps get 1 duty while prefects get 1 or 2). You may use a different value for this for each house - check with the House Captains whether they want year reps to get same number or less duties than prefects.

**dutiesThisWeekWeightingPB** - a value between 0 and 1 dictating how much ADA will avoid giving a prefect multiple duties in the same week. 0 means ADA will not look at the number of duties this week and allocate based on total number of duties, 1 means ADA will not look at the total duties and only make sure no one gets multiple duties in the week

**dutiesThisWeekWeightingHouse** - same as the previous value but for house prefects. I recommend making this higher than the previous since house prefects are less strict with making sure everyone has the same number of total duties and focus more on spreading out the duties across all prefects

The following settings require a more technical understanding of the algorithm and I recommend not modifying the values unless you know what you're doing. For a more complete picture of what these settings do, refer to the [prefect scoring algorithm](#prefect-scoring-algorithm)

**generalScoreWeighting** - a value between 0 and 1 dictating how much ADA will look at the number of total duties and duties this week compared to how well the specific duty matches. 0 means ADA will only consider the time since the last specific duty and number of specific duties, 1 means ADA will only consider the number of total duties and duties this week

**timeSinceSpecificDutyWeighting** - a value between 0 and 1 dictating how much ADA will look at the time since the last specific duty compared to the number of specific duties

**worseScoreWeighting** - a value between 0 and 1 dictating how much ADA will look at the worse of the general and specific scores comapred to the weighted average

# Technical details
## Prefect scoring algorithm

<img width="800" alt="Prefect Scoring Algorithm" src="https://github.com/user-attachments/assets/af23cd32-885e-4f5d-8e92-557cce3ecdf6" />


# Other notes
What to do when the subcommittee EXCOs are decided?
How to handle duty replacements?

How to set up probation duty weeks?
