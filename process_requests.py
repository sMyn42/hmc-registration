from threading import local
import numpy as np
import pandas as pd
import openpyxl as pxl
from itertools import islice
import os
import sys
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from utils import all_committees, team_committees, get_assignment, import_workbook, read_global_current, update_cap_sheet
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
  school_list = []
  # iterate over files in
  # that directory
  for filename in os.listdir(directory):
    f = os.path.join(directory, filename)
    # checking if it is a file
    if ".csv" in f:
      school_list.append(pd.read_csv("./" + f))
      # Move file to completed Directory
      os.rename(f, os.path.join(out_directory, filename))

  # do the stuff
  for school in school_list:
    school_shuffled = school.sample(frac=1)
    school_meta = school_shuffled.iloc[:, :6]
    school_prefs = school_shuffled.iloc[:, 6:14]
    local_current = dict(zip(all_committees, [0 for i in all_committees]))

    # print("GLOBALS")
    # print(global_current)
    # print("")
    # print("LOCALS")
    # print(local_current)
    # print("")

    # local caps refers to number of teams, multiplied by 6 for dc, 2 for others.
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


    # print("GLOBALS")
    # print(global_current)
    # print("")
    # print("LOCALS")
    # print(local_current)
    # print("")
    #update global sheet locations.
    update_cap_sheet(global_current, w["Main"])

    # Write each name (local_assn) to sheet.
    #print(local_assn)

    grouped_assn = local_assn.groupby("Committee")
    grouped_assn_dict = {x:grouped_assn.get_group(x) for x in grouped_assn.groups}

    #print(grouped_assn_dict)

    for c in all_committees:
      # print(grouped_assn_dict.get(c))
      if grouped_assn_dict.get(c) is not None:
        print("Printing " + c + " assignments to sheet.")
        for r in dataframe_to_rows(grouped_assn_dict.get(c), header=True):
          w[c].append(r)

    # write new worksheet to file
    # w["Main"] = new_ws
    w.save(out_file)

if __name__ == "__main__":
  main()