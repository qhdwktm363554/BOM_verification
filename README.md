# BOM_verification
BOM and program data (Siplace Mounter program) verification system

purpose:
  BOM from SAP(ERP) system and Placement list from Siplace(mounter) must be same based on its reference# and component part#

features:
  1. result file contains total placement Nr for OK and NG
  2. component recognition: regex (refer to the code inside)
  3. file format: .xls for BOM and .csv for placement list
