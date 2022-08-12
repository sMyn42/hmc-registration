from threading import local
import numpy as np
import pandas as pd
import openpyxl as pxl
from itertools import islice
import os
import sys
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from utils import all_committees, team_committees, get_assignment, import_workbook, read_global_current, update_cap_sheet, get_solo_assignment
from create_sheet import out_file
from itertools import islice


def main():

  # Assign w to workbook
  w, wr = import_workbook()

  # Read and create global current vars
  global_current = read_global_current(w["Main"])

  # Get school_list
  directory = 'pending-role-requests'
  out_directory = 'satisfied-role-requests'
  school_file_list = []
  # iterate over files in that directory

  for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a file
    if ".csv" in f or ".xlsx" in f:
      school_file_list.append(f)

  # do the stuff
  for f in school_file_list:
    school = None
    if ".csv" in f:
      school = pd.read_csv(f, dtype=object).fillna(" ")
    elif ".xlsx" in f:
      school = pd.read_excel(f, dtype=object, usecols='A:O').fillna(" ")
    school_shuffled = school.sample(frac=1)
    school_meta = school_shuffled.iloc[:, :6]
    school_prefs = school_shuffled.iloc[:, 6:14]
    local_current = dict(zip(all_committees, [0 for i in all_committees]))

    # local caps refers to number of teams, multiplied by 6 for dc, 2 for others.

    ########
    #TODO ARE TEAM LOCAL CAPS DONE RIGHT????
    #######

    team_local_caps = dict(zip(team_committees, [2, 2, 6]))
    local_caps = dict(zip(all_committees, [2 * team_local_caps[i] if i in team_committees else 4 for i in all_committees]))
    local_assn = pd.DataFrame(columns=["School", "First", "Last", "Grade", "Experience", "Email", "Committee"])

    for entry_num in range(school_shuffled.shape[0]):
      p = get_assignment(school_prefs.iloc[entry_num], {"local_caps": local_caps, "local_current":local_current})
      p_meta = school_meta.iloc[entry_num]
      p_meta = list(p_meta)
      p_meta.append(p)

      local_assn = pd.concat([local_assn, pd.DataFrame([p_meta], columns=local_assn.columns)])
      #update local caps
      local_current[p] = local_current[p] + 1
      #update global caps
      global_current[p] = global_current[p] + 1

    grouped_assn = local_assn.groupby("Committee")
    grouped_assn_dict = {x:grouped_assn.get_group(x) for x in grouped_assn.groups}

    #################
    # TODO TEAM COMMITTEE STUFF:
    #   1. CHECK local_assn FOR ANY ODD NUMBERED PPL IN TEAM COMMITTEES (based on exact committee name.)

    stragglers = pd.DataFrame(columns=school_shuffled.columns)

    for c, c_assn in grouped_assn_dict.items():
      if c in team_committees:
        team_n = team_local_caps[c]
        if c_assn.shape[0] % team_n != 0:
          n_teams = c_assn.shape[0] // team_n
          odds_names = c_assn.iloc[n_teams * team_n:, :]
          c_assn = c_assn.iloc[:n_teams * team_n, :]
          grouped_assn_dict[c] = c_assn
          for i in range(odds_names.shape[0]):
            print("REMOVING OVERFILLED TEAM COMMITTEE MEMBER!!!")
            local_current[c] = local_current[c] - 1
            global_current[c] = global_current[c] - 1
            prsn = odds_names.iloc[i]
            prsn_prefs = school_shuffled.loc[(school_shuffled["Delegate First Name"] == prsn["First"]) & (school_shuffled["Delegate Last Name"] == prsn["Last"])]
            stragglers = pd.concat([stragglers, prsn_prefs])

    #   2. CUT THEM OUT OF local_assn // UPDATE or DECREMENT caps (global and local) -> DO B4 2nd part. -> can add after team_assgnment part of get_assignment. 
    #         wait nah this don't work u don't have access to local_assn from the get_assignment function; also, you need to know when all team_assignments are done; 
    #         this only happens in control flow. Need to create another utils function that *takes a name* and only gives it a *solo committee*. -> 
    #         also delete from local_assn in main and add the new ret val to local_assn.
    #   3. ADD THEM BACK USING THE NEW UTILS FUNCTION -> SIMILAR TO MISC PART OF THE SECOND LOOP IN get_assnment.

    # 2nd assignment iteration:
    stragglers_meta = stragglers.iloc[:, :6]
    stragglers_prefs = stragglers.iloc[:, 6:14]

    for entry_num in range(stragglers.shape[0]):
      p = get_solo_assignment(stragglers_prefs.iloc[entry_num], {"local_caps": local_caps, "local_current":local_current})
      p_meta = stragglers_meta.iloc[entry_num]
      p_meta = list(stragglers_meta)
      p_meta.append(p)

      local_assn = pd.concat([local_assn, pd.DataFrame([p_meta], columns=local_assn.columns)])
      #update local caps
      local_current[p] = local_current[p] + 1
      #update global caps
      global_current[p] = global_current[p] + 1

    update_cap_sheet(global_current, w["Main"])

    ####################

    for c in all_committees:
      # print(grouped_assn_dict.get(c))
      if grouped_assn_dict.get(c) is not None:
        print("Printing " + c + " assignments to sheet.")
        for r in dataframe_to_rows(grouped_assn_dict.get(c), index=False, header=False):
          w[c].append(r)

    w.save(out_file)

    # Move file
    os.rename(f, f.replace(directory, out_directory))

if __name__ == "__main__":
  main()