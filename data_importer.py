from PyQt5 import QtWidgets
import pandas as pd
import os
import glob
from umts_monthly_report import UMTSReportGenerator
from lte_monthly_report import LTEReportGenerator

def conv_number_str(val):
    if val == 'NIL' or val == '':
        return '0'
    return str(val)

# def conv_int(val):
#     try:
#         return int(val)
#     except ValueError:
#         if val != 'NIL' and val != '':
#             print('ERROR')
#             print(type(val))
#             print(val)
#             raise ValueError('Converting to Int is losing data. Check the import files.')
#         return 0
    
# def conv_float(val):
#     try:
#         return float(val)
#     except ValueError:
#         if val != 'NIL' and val != '':
#             print('ERROR')
#             print(type(val))
#             print(val)
#             raise ValueError('Converting to Float is losing data. Check the import files.')
#         return 0

def conv_rtwp(val):
    try:
        return float(val)
    except ValueError:
        return -102 # Avg of RTWP

def conv_tcp(val):
    try:
        return float(val)
    except ValueError:
        return 39 # Avg of TCP

def conv_interference(val):
    try:
        return float(val)
    except ValueError:
        return -108 # Avg of interf

class CountersDataImporter():
    def umts_data_import1(self):
        self.umts_path_1 = QtWidgets.QFileDialog.getExistingDirectory(None,'Import UMTS Query 1',"F:\ ")
    
    def umts_data_import2(self):
        self.umts_path_2 = QtWidgets.QFileDialog.getExistingDirectory(None,'Import UMTS Query 2',"F:\ ")

    def umts_report_run(self):
        if hasattr(self, 'umtsCounters1_df') and hasattr(self, 'umtsCounters2_df'):
            pass
        else:
            # Query 1
            excel_files = glob.glob(os.path.join(self.umts_path_1, "*.csv"))
            appended_data = []
            print('Importing UMTS Queries (1)...')
            for f in excel_files:
                df = pd.read_csv(f, skiprows=7, converters={
                    'Start Time': str,
                    'NE Name': str,
                    'VS.MultRAB.SF128 (None)': conv_number_str,
                    'VS.MultRAB.SF16 (None)': conv_number_str,
                    'VS.MultRAB.SF256 (None)': conv_number_str,
                    'VS.MultRAB.SF32 (None)': conv_number_str,
                    'VS.MultRAB.SF4 (None)': conv_number_str,
                    'VS.MultRAB.SF64 (None)': conv_number_str,
                    'VS.MultRAB.SF8 (None)': conv_number_str,
                    'VS.SingleRAB.SF128 (None)': conv_number_str,
                    'VS.SingleRAB.SF16 (None)': conv_number_str,
                    'VS.SingleRAB.SF256 (None)': conv_number_str,
                    'VS.SingleRAB.SF32 (None)': conv_number_str,
                    'VS.SingleRAB.SF4 (None)': conv_number_str,
                    'VS.SingleRAB.SF64 (None)': conv_number_str,
                    'VS.SingleRAB.SF8 (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgConvCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgStrCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgInterCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgBkgCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgSubCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.TmConvCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.TmStrCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.TmItrCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.TmBkgCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.EmgCall (None)': conv_number_str,
                    'RRC.SuccConnEstab.Unkown (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgHhPrSig (None)': conv_number_str,
                    'RRC.SuccConnEstab.OrgLwPrSig (None)': conv_number_str,
                    'RRC.SuccConnEstab.CallReEst (None)': conv_number_str,
                    'RRC.SuccConnEstab.TmHhPrSig (None)': conv_number_str,
                    'RRC.SuccConnEstab.TmLwPrSig (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgConvCall (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgInterCall (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgStrCall (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgBkgCall (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgSubCall (None)': conv_number_str,
                    'RRC.AttConnEstab.TmBkgCall (None)': conv_number_str,
                    'RRC.AttConnEstab.TmConvCall (None)': conv_number_str,
                    'RRC.AttConnEstab.TmInterCall (None)': conv_number_str,
                    'RRC.AttConnEstab.TmStrCall (None)': conv_number_str,
                    'RRC.AttConnEstab.EmgCall (None)': conv_number_str,
                    'RRC.AttConnEstab.Unknown (None)': conv_number_str,
                    'RRC.AttConnEstab.CallReEst (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgHhPrSig (None)': conv_number_str,
                    'RRC.AttConnEstab.OrgLwPrSig (None)': conv_number_str,
                    'RRC.AttConnEstab.TmHhPrSig (None)': conv_number_str,
                    'RRC.AttConnEstab.TmLwPrSig (None)': conv_number_str,
                    'VS.SuccCellUpdt.PageRsp (None)': conv_number_str,
                    'VS.SuccCellUpdt.ULDataTrans (None)': conv_number_str,
                    'VS.SuccCellUpdt.Reg.PCH (None)': conv_number_str,
                    'VS.SuccCellUpdt.Detach.PCH (None)': conv_number_str,
                    'VS.SuccCellUpdt.Other.PCH (None)': conv_number_str,
                    'VS.AttCellUpdt.PageRsp (None)': conv_number_str,
                    'VS.AttCellUpdt.ULDataTrans (None)': conv_number_str,
                    'VS.AttCellUpdt.Reg.PCH (None)': conv_number_str,
                    'VS.AttCellUpdt.Detach.PCH (None)': conv_number_str,
                    'VS.AttCellUpdt.Other.PCH (None)': conv_number_str,
                    'VS.RRC.Rej.ULPower.Cong (None)': conv_number_str,
                    'VS.RRC.Rej.DLPower.Cong (None)': conv_number_str,
                    'VS.RRC.Rej.ULIUBBand.Cong (None)': conv_number_str,
                    'VS.RRC.Rej.DLIUBBand.Cong (None)': conv_number_str,
                    'VS.RRC.Rej.ULCE.Cong (None)': conv_number_str,
                    'VS.RRC.Rej.DLCE.Cong (None)': conv_number_str,
                    'VS.RRC.Rej.Code.Cong (None)': conv_number_str,
                    'VS.RRC.AttConnEstab.Sum (None)': conv_number_str,
                    'VS.RAB.SuccEstabCS.Conv (None)': conv_number_str,
                    'VS.RAB.SuccEstabCS.Str (None)': conv_number_str,
                    'VS.RAB.AttEstabCS.Conv (None)': conv_number_str,
                    'VS.RAB.AttEstabCS.Str (None)': conv_number_str,
                    'VS.RAB.SuccEstabPS.Conv (None)': conv_number_str,
                    'VS.RAB.SuccEstabPS.Str (None)': conv_number_str,
                    'VS.RAB.SuccEstabPS.Int (None)': conv_number_str,
                    'VS.RAB.SuccEstabPS.Bkg (None)': conv_number_str,
                    'VS.RAB.AttEstabPS.Conv (None)': conv_number_str,
                    'VS.RAB.AttEstabPS.Str (None)': conv_number_str,
                    'VS.RAB.AttEstabPS.Int (None)': conv_number_str,
                    'VS.RAB.AttEstabPS.Bkg (None)': conv_number_str,
                    'VS.RRC.Paging1.Loss.PCHCong.Cell (None)': conv_number_str,
                    'VS.UTRAN.AttPaging1 (None)': conv_number_str,
                    'VS.RAB.AbnormRel.CS (None)': conv_number_str,
                    'VS.RAB.NormRel.CS (None)': conv_number_str,
                    'VS.RAB.AbnormRel.PS (None)': conv_number_str,
                    'VS.RAB.NormRel.PS (None)': conv_number_str,
                    'VS.HSDPA.RAB.AbnormRel (None)': conv_number_str,
                    'VS.HSDPA.RAB.AbnormRel.H2P (None)': conv_number_str,
                    'VS.HSDPA.RAB.NormRel (None)': conv_number_str,
                    'VS.HSDPA.HHO.H2D.SuccOutIntraFreq (None)': conv_number_str,
                    'VS.HSDPA.HHO.H2D.SuccOutInterFreq (None)': conv_number_str,
                    'VS.HSDPA.H2D.Succ (None)': conv_number_str,
                    'VS.HSDPA.H2F.Succ (None)': conv_number_str,
                    'VS.HSDPA.H2P.Succ (None)': conv_number_str,
                    'VS.RAB.AbnormRel.PSR99 (None)': conv_number_str,
                    'VS.RAB.NormRel.PSR99 (None)': conv_number_str,
                    'VS.HSDPA.MeanChThroughput (kbit/s)': conv_number_str,
                    'VS.Cell.UnavailTime (s)': conv_number_str,
                    'VS.Cell.UnavailTime.Sys (s)': conv_number_str,
                    'VS.HSDPA.UE.Mean.Cell (None)': conv_number_str
                    })
                df.drop(['Period(min)'], axis=1, inplace=True)
                if 'BSC6910UCell' in df.columns:
                    df.rename(columns = {'BSC6910UCell':'Cell'}, inplace = True)
                elif 'BSC6900UCell' in df.columns:
                    df.rename(columns = {'BSC6900UCell':'Cell'}, inplace = True)
                appended_data.append(df)
                print('Imported UMTS query 1: '+f)
            if appended_data:
                self.umtsCounters1_df = pd.concat(appended_data)
                self.umtsCounters1_df = self.umtsCounters1_df.astype({
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'VS.MultRAB.SF128 (None)': 'float64',
                    'VS.MultRAB.SF16 (None)': 'float64',
                    'VS.MultRAB.SF256 (None)': 'float64',
                    'VS.MultRAB.SF32 (None)': 'float64',
                    'VS.MultRAB.SF4 (None)': 'float64',
                    'VS.MultRAB.SF64 (None)': 'float64',
                    'VS.MultRAB.SF8 (None)': 'float64',
                    'VS.SingleRAB.SF128 (None)': 'float64',
                    'VS.SingleRAB.SF16 (None)': 'float64',
                    'VS.SingleRAB.SF256 (None)': 'float64',
                    'VS.SingleRAB.SF32 (None)': 'float64',
                    'VS.SingleRAB.SF4 (None)': 'float64',
                    'VS.SingleRAB.SF64 (None)': 'float64',
                    'VS.SingleRAB.SF8 (None)': 'float64',
                    'RRC.SuccConnEstab.OrgConvCall (None)': 'int64',
                    'RRC.SuccConnEstab.OrgStrCall (None)': 'int64',
                    'RRC.SuccConnEstab.OrgInterCall (None)': 'int64',
                    'RRC.SuccConnEstab.OrgBkgCall (None)': 'int64',
                    'RRC.SuccConnEstab.OrgSubCall (None)': 'int64',
                    'RRC.SuccConnEstab.TmConvCall (None)': 'int64',
                    'RRC.SuccConnEstab.TmStrCall (None)': 'int64',
                    'RRC.SuccConnEstab.TmItrCall (None)': 'int64',
                    'RRC.SuccConnEstab.TmBkgCall (None)': 'int64',
                    'RRC.SuccConnEstab.EmgCall (None)': 'int64',
                    'RRC.SuccConnEstab.Unkown (None)': 'int64',
                    'RRC.SuccConnEstab.OrgHhPrSig (None)': 'int64',
                    'RRC.SuccConnEstab.OrgLwPrSig (None)': 'int64',
                    'RRC.SuccConnEstab.CallReEst (None)': 'int64',
                    'RRC.SuccConnEstab.TmHhPrSig (None)': 'int64',
                    'RRC.SuccConnEstab.TmLwPrSig (None)': 'int64',
                    'RRC.AttConnEstab.OrgConvCall (None)': 'int64',
                    'RRC.AttConnEstab.OrgInterCall (None)': 'int64',
                    'RRC.AttConnEstab.OrgStrCall (None)': 'int64',
                    'RRC.AttConnEstab.OrgBkgCall (None)': 'int64',
                    'RRC.AttConnEstab.OrgSubCall (None)': 'int64',
                    'RRC.AttConnEstab.TmBkgCall (None)': 'int64',
                    'RRC.AttConnEstab.TmConvCall (None)': 'int64',
                    'RRC.AttConnEstab.TmInterCall (None)': 'int64',
                    'RRC.AttConnEstab.TmStrCall (None)': 'int64',
                    'RRC.AttConnEstab.EmgCall (None)': 'int64',
                    'RRC.AttConnEstab.Unknown (None)': 'int64',
                    'RRC.AttConnEstab.CallReEst (None)': 'int64',
                    'RRC.AttConnEstab.OrgHhPrSig (None)': 'int64',
                    'RRC.AttConnEstab.OrgLwPrSig (None)': 'int64',
                    'RRC.AttConnEstab.TmHhPrSig (None)': 'int64',
                    'RRC.AttConnEstab.TmLwPrSig (None)': 'int64',
                    'VS.SuccCellUpdt.PageRsp (None)': 'int64',
                    'VS.SuccCellUpdt.ULDataTrans (None)': 'int64',
                    'VS.SuccCellUpdt.Reg.PCH (None)': 'int64',
                    'VS.SuccCellUpdt.Detach.PCH (None)': 'int64',
                    'VS.SuccCellUpdt.Other.PCH (None)': 'int64',
                    'VS.AttCellUpdt.PageRsp (None)': 'int64',
                    'VS.AttCellUpdt.ULDataTrans (None)': 'int64',
                    'VS.AttCellUpdt.Reg.PCH (None)': 'int64',
                    'VS.AttCellUpdt.Detach.PCH (None)': 'int64',
                    'VS.AttCellUpdt.Other.PCH (None)': 'int64',
                    'VS.RRC.Rej.ULPower.Cong (None)': 'int64',
                    'VS.RRC.Rej.DLPower.Cong (None)': 'int64',
                    'VS.RRC.Rej.ULIUBBand.Cong (None)': 'int64',
                    'VS.RRC.Rej.DLIUBBand.Cong (None)': 'int64',
                    'VS.RRC.Rej.ULCE.Cong (None)': 'int64',
                    'VS.RRC.Rej.DLCE.Cong (None)': 'int64',
                    'VS.RRC.Rej.Code.Cong (None)': 'int64',
                    'VS.RRC.AttConnEstab.Sum (None)': 'int64',
                    'VS.RAB.SuccEstabCS.Conv (None)': 'int64',
                    'VS.RAB.SuccEstabCS.Str (None)': 'int64',
                    'VS.RAB.AttEstabCS.Conv (None)': 'int64',
                    'VS.RAB.AttEstabCS.Str (None)': 'int64',
                    'VS.RAB.SuccEstabPS.Conv (None)': 'int64',
                    'VS.RAB.SuccEstabPS.Str (None)': 'int64',
                    'VS.RAB.SuccEstabPS.Int (None)': 'int64',
                    'VS.RAB.SuccEstabPS.Bkg (None)': 'int64',
                    'VS.RAB.AttEstabPS.Conv (None)': 'int64',
                    'VS.RAB.AttEstabPS.Str (None)': 'int64',
                    'VS.RAB.AttEstabPS.Int (None)': 'int64',
                    'VS.RAB.AttEstabPS.Bkg (None)': 'int64',
                    'VS.RRC.Paging1.Loss.PCHCong.Cell (None)': 'int64',
                    'VS.UTRAN.AttPaging1 (None)': 'int64',
                    'VS.RAB.AbnormRel.CS (None)': 'int64',
                    'VS.RAB.NormRel.CS (None)': 'int64',
                    'VS.RAB.AbnormRel.PS (None)': 'int64',
                    'VS.RAB.NormRel.PS (None)': 'int64',
                    'VS.HSDPA.RAB.AbnormRel (None)': 'int64',
                    'VS.HSDPA.RAB.AbnormRel.H2P (None)': 'int64',
                    'VS.HSDPA.RAB.NormRel (None)': 'int64',
                    'VS.HSDPA.HHO.H2D.SuccOutIntraFreq (None)': 'int64',
                    'VS.HSDPA.HHO.H2D.SuccOutInterFreq (None)': 'int64',
                    'VS.HSDPA.H2D.Succ (None)': 'int64',
                    'VS.HSDPA.H2F.Succ (None)': 'int64',
                    'VS.HSDPA.H2P.Succ (None)': 'int64',
                    'VS.RAB.AbnormRel.PSR99 (None)': 'int64',
                    'VS.RAB.NormRel.PSR99 (None)': 'int64',
                    'VS.HSDPA.MeanChThroughput (kbit/s)': 'float64',
                    'VS.Cell.UnavailTime (s)': 'int64',
                    'VS.Cell.UnavailTime.Sys (s)': 'int64',
                    'VS.HSDPA.UE.Mean.Cell (None)': 'float64'
                })
                self.umtsCounters1_df = self.umtsCounters1_df.drop_duplicates(subset=['Start Time', 'Cell'])
                print('UMTS Queries 1 imported successfully')
            # Query 2
            excel_files = glob.glob(os.path.join(self.umts_path_2, "*.csv"))
            appended_data = []
            print('Importing UMTS Queries (2)...')
            for f in excel_files:
                df = pd.read_csv(f, skiprows=7, converters={
                    'Start Time': str,
                    'NE Name': str,
                    'VS.MeanTCP (dBm)': conv_tcp,
                    'VS.MeanTCP.NonHS (dBm)': conv_tcp,
                    'VS.HSUPA.RAB.AbnormRel (None)': conv_number_str,
                    'VS.HSUPA.RAB.AbnormRel.E2P (None)': conv_number_str,
                    'VS.HSUPA.RAB.NormRel (None)': conv_number_str,
                    'VS.HSUPA.HHO.E2D.SuccOutIntraFreq (None)': conv_number_str,
                    'VS.HSUPA.HHO.E2D.SuccOutInterFreq (None)': conv_number_str,
                    'VS.HSUPA.E2F.Succ (None)': conv_number_str,
                    'VS.HSUPA.E2D.Succ (None)': conv_number_str,
                    'VS.HSUPA.E2P.Succ (None)': conv_number_str,
                    'VS.SHO.SuccRLAdd (None)': conv_number_str,
                    'VS.SHO.SuccRLDel (None)': conv_number_str,
                    'VS.SHO.AttRLAdd (None)': conv_number_str,
                    'VS.SHO.AttRLDel (None)': conv_number_str,
                    'VS.HHO.SuccIntraFreqOut.IntraNodeB (None)': conv_number_str,
                    'VS.HHO.SuccIntraFreqOut.InterNodeBIntraRNC (None)': conv_number_str,
                    'VS.HHO.SuccIntraFreqOut.InterRNC (None)': conv_number_str,
                    'VS.HHO.AttIntraFreqOut.InterNodeBIntraRNC (None)': conv_number_str,
                    'VS.HHO.AttIntraFreqOut.InterRNC (None)': conv_number_str,
                    'VS.HHO.AttIntraFreqOut.IntraNodeB (None)': conv_number_str,
                    'VS.HHO.SuccInterFreqOut (None)': conv_number_str,
                    'VS.HHO.AttInterFreqOut (None)': conv_number_str,
                    'VS.HSUPA.MeanChThroughput (kbit/s)': conv_number_str,
                    'VS.MeanRTWP (dBm)': conv_rtwp,
                    'VS.MaxRTWP (dBm)': conv_rtwp,
                    'VS.MinRTWP (dBm)': conv_rtwp,
                    'VS.RAB.AMR.Erlang.cell (Erl)': conv_number_str,
                    'VS.RAB.AMRWB.Erlang.cell (Erl)': conv_number_str,
                    'VS.PS.Bkg.DL.128.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.144.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.16.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.256.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.32.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.384.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.64.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.DL.8.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.128.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.144.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.16.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.256.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.32.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.384.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.64.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.DL.8.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.128.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.144.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.16.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.256.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.32.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.384.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.64.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.DL.8.Traffic (bit)': conv_number_str,
                    'VS.PS.Conv.DL.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.128.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.144.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.16.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.256.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.32.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.384.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.64.Traffic (bit)': conv_number_str,
                    'VS.PS.Bkg.UL.8.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.128.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.144.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.16.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.256.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.32.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.384.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.64.Traffic (bit)': conv_number_str,
                    'VS.PS.Int.UL.8.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.UL.128.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.UL.16.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.UL.32.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.UL.64.Traffic (bit)': conv_number_str,
                    'VS.PS.Str.UL.8.Traffic (bit)': conv_number_str,
                    'VS.PS.Conv.UL.Traffic (bit)': conv_number_str,
                    'VS.HSDPA.MeanChThroughput.TotalBytes (byte)': conv_number_str,
                    'VS.HSUPA.MeanChThroughput.TotalBytes (byte)': conv_number_str,
                    'VS.CellDCHUEs (None)': conv_number_str,
                    'VS.HSUPA.UE.Mean.Cell (None)': conv_number_str,
                    'VS.PSLoad.ULThruput.MOCN.PLMN0 (byte)': conv_number_str,
                    'VS.PSLoad.ULThruput.MOCN.PLMN1 (byte)': conv_number_str,
                    'VS.PSLoad.ULThruput.MOCN.PLMN2 (byte)': conv_number_str,
                    'VS.PSLoad.ULThruput.MOCN.PLMN3 (byte)': conv_number_str,
                    'VS.PSLoad.DLThruput.MOCN.PLMN0 (byte)': conv_number_str,
                    'VS.PSLoad.DLThruput.MOCN.PLMN1 (byte)': conv_number_str,
                    'VS.PSLoad.DLThruput.MOCN.PLMN2 (byte)': conv_number_str,
                    'VS.PSLoad.DLThruput.MOCN.PLMN3 (byte)': conv_number_str,
                    'VS.CS.Erlang.Equiv.MOCN.PLMN0 (Erl)': conv_number_str,
                    'VS.CS.Erlang.Equiv.MOCN.PLMN1 (Erl)': conv_number_str,
                    'VS.CS.Erlang.Equiv.MOCN.PLMN2 (Erl)': conv_number_str,
                    'VS.CS.Erlang.Equiv.MOCN.PLMN3 (Erl)': conv_number_str       
                    })
                df.drop(['Period(min)'], axis=1, inplace=True)
                if 'BSC6910UCell' in df.columns:
                    df.rename(columns = {'BSC6910UCell':'Cell'}, inplace = True)
                elif 'BSC6900UCell' in df.columns:
                    df.rename(columns = {'BSC6900UCell':'Cell'}, inplace = True)
                appended_data.append(df)
                print('Imported UMTS query 2: '+f)
            if appended_data:
                self.umtsCounters2_df = pd.concat(appended_data)
                self.umtsCounters2_df = self.umtsCounters2_df.astype({
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'VS.MeanTCP (dBm)': 'float64',
                    'VS.MeanTCP.NonHS (dBm)': 'float64',
                    'VS.HSUPA.RAB.AbnormRel (None)': 'int64',
                    'VS.HSUPA.RAB.AbnormRel.E2P (None)': 'int64',
                    'VS.HSUPA.RAB.NormRel (None)': 'int64',
                    'VS.HSUPA.HHO.E2D.SuccOutIntraFreq (None)': 'int64',
                    'VS.HSUPA.HHO.E2D.SuccOutInterFreq (None)': 'int64',
                    'VS.HSUPA.E2F.Succ (None)': 'int64',
                    'VS.HSUPA.E2D.Succ (None)': 'int64',
                    'VS.HSUPA.E2P.Succ (None)': 'int64',
                    'VS.SHO.SuccRLAdd (None)': 'int64',
                    'VS.SHO.SuccRLDel (None)': 'int64',
                    'VS.SHO.AttRLAdd (None)': 'int64',
                    'VS.SHO.AttRLDel (None)': 'int64',
                    'VS.HHO.SuccIntraFreqOut.IntraNodeB (None)': 'int64',
                    'VS.HHO.SuccIntraFreqOut.InterNodeBIntraRNC (None)': 'int64',
                    'VS.HHO.SuccIntraFreqOut.InterRNC (None)': 'int64',
                    'VS.HHO.AttIntraFreqOut.InterNodeBIntraRNC (None)': 'int64',
                    'VS.HHO.AttIntraFreqOut.InterRNC (None)': 'int64',
                    'VS.HHO.AttIntraFreqOut.IntraNodeB (None)': 'int64',
                    'VS.HHO.SuccInterFreqOut (None)': 'int64',
                    'VS.HHO.AttInterFreqOut (None)': 'int64',
                    'VS.HSUPA.MeanChThroughput (kbit/s)': 'float64',
                    'VS.MeanRTWP (dBm)': 'float64',
                    'VS.MaxRTWP (dBm)': 'float64',
                    'VS.MinRTWP (dBm)': 'float64',
                    'VS.RAB.AMR.Erlang.cell (Erl)': 'float64',
                    'VS.RAB.AMRWB.Erlang.cell (Erl)': 'float64',
                    'VS.PS.Bkg.DL.128.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.144.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.16.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.256.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.32.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.384.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.64.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.DL.8.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.128.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.144.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.16.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.256.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.32.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.384.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.64.Traffic (bit)': 'int64',
                    'VS.PS.Int.DL.8.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.128.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.144.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.16.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.256.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.32.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.384.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.64.Traffic (bit)': 'int64',
                    'VS.PS.Str.DL.8.Traffic (bit)': 'int64',
                    'VS.PS.Conv.DL.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.128.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.144.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.16.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.256.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.32.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.384.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.64.Traffic (bit)': 'int64',
                    'VS.PS.Bkg.UL.8.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.128.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.144.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.16.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.256.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.32.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.384.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.64.Traffic (bit)': 'int64',
                    'VS.PS.Int.UL.8.Traffic (bit)': 'int64',
                    'VS.PS.Str.UL.128.Traffic (bit)': 'int64',
                    'VS.PS.Str.UL.16.Traffic (bit)': 'int64',
                    'VS.PS.Str.UL.32.Traffic (bit)': 'int64',
                    'VS.PS.Str.UL.64.Traffic (bit)': 'int64',
                    'VS.PS.Str.UL.8.Traffic (bit)': 'int64',
                    'VS.PS.Conv.UL.Traffic (bit)': 'int64',
                    'VS.HSDPA.MeanChThroughput.TotalBytes (byte)': 'int64',
                    'VS.HSUPA.MeanChThroughput.TotalBytes (byte)': 'int64',
                    'VS.CellDCHUEs (None)': 'float64',
                    'VS.HSUPA.UE.Mean.Cell (None)': 'float64',
                    'VS.PSLoad.ULThruput.MOCN.PLMN0 (byte)': 'int64',
                    'VS.PSLoad.ULThruput.MOCN.PLMN1 (byte)': 'int64',
                    'VS.PSLoad.ULThruput.MOCN.PLMN2 (byte)': 'int64',
                    'VS.PSLoad.ULThruput.MOCN.PLMN3 (byte)': 'int64',
                    'VS.PSLoad.DLThruput.MOCN.PLMN0 (byte)': 'int64',
                    'VS.PSLoad.DLThruput.MOCN.PLMN1 (byte)': 'int64',
                    'VS.PSLoad.DLThruput.MOCN.PLMN2 (byte)': 'int64',
                    'VS.PSLoad.DLThruput.MOCN.PLMN3 (byte)': 'int64',
                    'VS.CS.Erlang.Equiv.MOCN.PLMN0 (Erl)': 'float64',
                    'VS.CS.Erlang.Equiv.MOCN.PLMN1 (Erl)': 'float64',
                    'VS.CS.Erlang.Equiv.MOCN.PLMN2 (Erl)': 'float64',
                    'VS.CS.Erlang.Equiv.MOCN.PLMN3 (Erl)': 'float64' 
                })
                self.umtsCounters2_df = self.umtsCounters2_df.drop_duplicates(subset=['Start Time', 'Cell'])
                print('UMTS Queries 2 imported successfully')
            umts_kpi = UMTSReportGenerator()
            umts_kpi.generate_report_data(self.umtsCounters1_df, self.umtsCounters2_df)

    def lte_data_import1(self):
        self.lte_path_1 = QtWidgets.QFileDialog.getExistingDirectory(None,'Import LTE Query 1',"F:\ ")
    
    def lte_data_import2(self):
        self.lte_path_2 = QtWidgets.QFileDialog.getExistingDirectory(None,'Import LTE Query 2',"F:\ ")
    
    def lte_data_import_mocn(self):
        self.lte_path_mocn = QtWidgets.QFileDialog.getExistingDirectory(None,'Import LTE Query MOCN',"F:\ ")
    
    def lte_report_run(self):
        if hasattr(self, 'lteCounters1_df') and hasattr(self, 'lteCounters2_df') and hasattr(self, 'lteCountersMocn_df'):
            pass
        else:
            # Query 1
            excel_files = glob.glob(os.path.join(self.lte_path_1, "*.csv"))
            appended_data = []
            print('Importing LTE Queries (1)...')
            for f in excel_files:
                df = pd.read_csv(f, skiprows=7, converters={
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'L.Thrp.bits.DL (bit)': conv_number_str,
                    'L.Thrp.Time.Cell.DL.HighPrecision (ms)': conv_number_str,
                    'L.Thrp.bits.UL (bit)': conv_number_str,
                    'L.Thrp.Time.Cell.UL.HighPrecision (ms)': conv_number_str,
                    'L.ChMeas.PRB.DL.Used.Avg (None)': conv_number_str,
                    'L.ChMeas.PRB.DL.Avail (None)': conv_number_str,
                    'L.ChMeas.PRB.UL.Used.Avg (None)': conv_number_str,
                    'L.ChMeas.PRB.UL.Avail (None)': conv_number_str,
                    'L.Traffic.ActiveUser.Avg (None)': conv_number_str,
                    'L.Traffic.User.Avg (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.MoSig (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.MoSig (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.DelayTol (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.Emc (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.HighPri (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.MoData (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.MoVoiceCall (None)': conv_number_str,
                    'L.RRC.ConnReq.Succ.Mt (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.DelayTol (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.Emc (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.HighPri (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.MoData (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.MoVoiceCall (None)': conv_number_str,
                    'L.RRC.ConnReq.Att.Mt (None)': conv_number_str,
                    'L.E-RAB.SuccEst.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.AttEst.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.FailEst.X2AP.VoIP (None)': conv_number_str,
                    'L.E-RAB.SuccEst (None)': conv_number_str,
                    'L.E-RAB.AttEst (None)': conv_number_str,
                    'L.E-RAB.FailEst.X2AP (None)': conv_number_str,
                    'L.S1Sig.ConnEst.Succ (None)': conv_number_str,
                    'L.S1Sig.ConnEst.Att (None)': conv_number_str,
                    'L.E-RAB.AbnormRel (None)': conv_number_str,
                    'L.E-RAB.NormRel (None)': conv_number_str,
                    'L.E-RAB.NormRel.IRatHOOut (None)': conv_number_str,
                    'L.E-RAB.AbnormRel.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.NormRel.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)': conv_number_str,
                    'L.HHO.IntraeNB.IntraFreq.ExecSuccOut (None)': conv_number_str,
                    'L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)': conv_number_str,
                    'L.HHO.IntraeNB.IntraFreq.ExecAttOut (None)': conv_number_str,
                    'L.HHO.IntereNB.IntraFreq.ExecAttOut (None)': conv_number_str,
                    'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntereNB.IntraFreq.ExecSuccOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntraeNB.IntraFreq.ExecAttOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntereNB.IntraFreq.ExecAttOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntraeNB.InterFreq.ExecSuccOut (None)': conv_number_str,
                    'L.HHO.IntereNB.InterFreq.ExecSuccOut (None)': conv_number_str,
                    'L.HHO.IntraeNB.InterFreq.ExecAttOut (None)': conv_number_str,
                    'L.HHO.IntereNB.InterFreq.ExecAttOut (None)': conv_number_str,
                    'L.HHO.IntraeNB.InterFreq.ExecSuccOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntereNB.InterFreq.ExecSuccOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntraeNB.InterFreq.ExecAttOut.VoIP (None)': conv_number_str,
                    'L.HHO.IntereNB.InterFreq.ExecAttOut.VoIP (None)': conv_number_str,
                    'L.CSFB.PrepSucc (None)': conv_number_str,
                    'L.CSFB.PrepAtt (None)': conv_number_str,
                    'L.CSFB.E2W (None)': conv_number_str,
                    'L.IRATHO.SRVCC.E2W.ExecSuccOut (None)': conv_number_str,
                    'L.IRATHO.SRVCC.E2W.MMEAbnormRsp (None)': conv_number_str,
                    'L.IRATHO.SRVCC.E2W.ExecAttOut (None)': conv_number_str,
                    'L.Thrp.bits.DL.LastTTI (bit)': conv_number_str,
                    'L.Thrp.Time.DL.RmvLastTTI (ms)': conv_number_str,
                    'L.Thrp.bits.UE.UL.LastTTI (bit)': conv_number_str,
                    'L.Thrp.Time.UE.UL.RmvLastTTI (ms)': conv_number_str,
                    'L.Thrp.bits.DL.QCI.9 (bit)': conv_number_str,
                    'L.Thrp.bits.DL.LastTTI.QCI.9 (bit)': conv_number_str,
                    'L.Thrp.Time.DL.RmvLastTTI.QCI.9 (ms)': conv_number_str,
                    'L.Thrp.bits.UL.QCI.9 (bit)': conv_number_str,
                    'L.Thrp.Time.UL.QCI.9 (ms)': conv_number_str,
                    'L.Thrp.bits.DL.QCI.8 (bit)': conv_number_str,
                    'L.Thrp.bits.DL.LastTTI.QCI.8 (bit)': conv_number_str,
                    'L.Thrp.Time.DL.RmvLastTTI.QCI.8 (ms)': conv_number_str,
                    'L.Thrp.bits.UL.QCI.8 (bit)': conv_number_str,
                    'L.Thrp.Time.UL.QCI.8 (ms)': conv_number_str,
                    'L.Traffic.DL.PktUuLoss.Loss (packet)': conv_number_str,
                    'L.Traffic.DL.PktUuLoss.Tot (packet)': conv_number_str,
                    'L.Traffic.DL.PktUuLoss.Loss.QCI.1 (packet)': conv_number_str,
                    'L.Traffic.DL.PktUuLoss.Tot.QCI.1 (packet)': conv_number_str,
                    'L.Traffic.UL.PktLoss.Loss (packet)': conv_number_str,
                    'L.Traffic.UL.PktLoss.Tot (packet)': conv_number_str,
                    'L.Traffic.UL.PktLoss.Loss.QCI.1 (packet)': conv_number_str,
                    'L.Traffic.UL.PktLoss.Tot.QCI.1 (packet)': conv_number_str,
                    'L.E-RAB.SessionTime.HighPrecision.QCI1 (100 ms)': conv_number_str
                })
                df.drop(['Period(min)'], axis=1, inplace=True)
                appended_data.append(df)
                print('Imported LTE query 1: '+f)
            if appended_data:
                self.lteCounters1_df = pd.concat(appended_data)
                self.lteCounters1_df = self.lteCounters1_df.astype({
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'L.Thrp.bits.DL (bit)': 'float64',
                    'L.Thrp.Time.Cell.DL.HighPrecision (ms)': 'int64',
                    'L.Thrp.bits.UL (bit)': 'float64',
                    'L.Thrp.Time.Cell.UL.HighPrecision (ms)': 'int64',
                    'L.ChMeas.PRB.DL.Used.Avg (None)': 'float64',
                    'L.ChMeas.PRB.DL.Avail (None)': 'int64',
                    'L.ChMeas.PRB.UL.Used.Avg (None)': 'float64',
                    'L.ChMeas.PRB.UL.Avail (None)': 'int64',
                    'L.Traffic.ActiveUser.Avg (None)': 'float64',
                    'L.Traffic.User.Avg (None)': 'float64',
                    'L.RRC.ConnReq.Succ.MoSig (None)': 'int64',
                    'L.RRC.ConnReq.Att.MoSig (None)': 'int64',
                    'L.RRC.ConnReq.Succ.DelayTol (None)': 'int64',
                    'L.RRC.ConnReq.Succ.Emc (None)': 'int64',
                    'L.RRC.ConnReq.Succ.HighPri (None)': 'int64',
                    'L.RRC.ConnReq.Succ.MoData (None)': 'int64',
                    'L.RRC.ConnReq.Succ.MoVoiceCall (None)': 'int64',
                    'L.RRC.ConnReq.Succ.Mt (None)': 'int64',
                    'L.RRC.ConnReq.Att.DelayTol (None)': 'int64',
                    'L.RRC.ConnReq.Att.Emc (None)': 'int64',
                    'L.RRC.ConnReq.Att.HighPri (None)': 'int64',
                    'L.RRC.ConnReq.Att.MoData (None)': 'int64',
                    'L.RRC.ConnReq.Att.MoVoiceCall (None)': 'int64',
                    'L.RRC.ConnReq.Att.Mt (None)': 'int64',
                    'L.E-RAB.SuccEst.QCI.1 (None)': 'int64',
                    'L.E-RAB.AttEst.QCI.1 (None)': 'int64',
                    'L.E-RAB.FailEst.X2AP.VoIP (None)': 'int64',
                    'L.E-RAB.SuccEst (None)': 'int64',
                    'L.E-RAB.AttEst (None)': 'int64',
                    'L.E-RAB.FailEst.X2AP (None)': 'int64',
                    'L.S1Sig.ConnEst.Succ (None)': 'int64',
                    'L.S1Sig.ConnEst.Att (None)': 'int64',
                    'L.E-RAB.AbnormRel (None)': 'int64',
                    'L.E-RAB.NormRel (None)': 'int64',
                    'L.E-RAB.NormRel.IRatHOOut (None)': 'int64',
                    'L.E-RAB.AbnormRel.QCI.1 (None)': 'int64',
                    'L.E-RAB.NormRel.QCI.1 (None)': 'int64',
                    'L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)': 'int64',
                    'L.HHO.IntraeNB.IntraFreq.ExecSuccOut (None)': 'int64',
                    'L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)': 'int64',
                    'L.HHO.IntraeNB.IntraFreq.ExecAttOut (None)': 'int64',
                    'L.HHO.IntereNB.IntraFreq.ExecAttOut (None)': 'int64',
                    'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.VoIP (None)': 'int64',
                    'L.HHO.IntereNB.IntraFreq.ExecSuccOut.VoIP (None)': 'int64',
                    'L.HHO.IntraeNB.IntraFreq.ExecAttOut.VoIP (None)': 'int64',
                    'L.HHO.IntereNB.IntraFreq.ExecAttOut.VoIP (None)': 'int64',
                    'L.HHO.IntraeNB.InterFreq.ExecSuccOut (None)': 'int64',
                    'L.HHO.IntereNB.InterFreq.ExecSuccOut (None)': 'int64',
                    'L.HHO.IntraeNB.InterFreq.ExecAttOut (None)': 'int64',
                    'L.HHO.IntereNB.InterFreq.ExecAttOut (None)': 'int64',
                    'L.HHO.IntraeNB.InterFreq.ExecSuccOut.VoIP (None)': 'int64',
                    'L.HHO.IntereNB.InterFreq.ExecSuccOut.VoIP (None)': 'int64',
                    'L.HHO.IntraeNB.InterFreq.ExecAttOut.VoIP (None)': 'int64',
                    'L.HHO.IntereNB.InterFreq.ExecAttOut.VoIP (None)': 'int64',
                    'L.CSFB.PrepSucc (None)': 'int64',
                    'L.CSFB.PrepAtt (None)': 'int64',
                    'L.CSFB.E2W (None)': 'int64',
                    'L.IRATHO.SRVCC.E2W.ExecSuccOut (None)': 'int64',
                    'L.IRATHO.SRVCC.E2W.MMEAbnormRsp (None)': 'int64',
                    'L.IRATHO.SRVCC.E2W.ExecAttOut (None)': 'int64',
                    'L.Thrp.bits.DL.LastTTI (bit)': 'float64',
                    'L.Thrp.Time.DL.RmvLastTTI (ms)': 'int64',
                    'L.Thrp.bits.UE.UL.LastTTI (bit)': 'float64',
                    'L.Thrp.Time.UE.UL.RmvLastTTI (ms)': 'int64',
                    'L.Thrp.bits.DL.QCI.9 (bit)': 'float64',
                    'L.Thrp.bits.DL.LastTTI.QCI.9 (bit)': 'float64',
                    'L.Thrp.Time.DL.RmvLastTTI.QCI.9 (ms)': 'int64',
                    'L.Thrp.bits.UL.QCI.9 (bit)': 'float64',
                    'L.Thrp.Time.UL.QCI.9 (ms)': 'int64',
                    'L.Thrp.bits.DL.QCI.8 (bit)': 'float64',
                    'L.Thrp.bits.DL.LastTTI.QCI.8 (bit)': 'float64',
                    'L.Thrp.Time.DL.RmvLastTTI.QCI.8 (ms)': 'int64',
                    'L.Thrp.bits.UL.QCI.8 (bit)': 'float64',
                    'L.Thrp.Time.UL.QCI.8 (ms)': 'int64',
                    'L.Traffic.DL.PktUuLoss.Loss (packet)': 'int64',
                    'L.Traffic.DL.PktUuLoss.Tot (packet)': 'int64',
                    'L.Traffic.DL.PktUuLoss.Loss.QCI.1 (packet)': 'int64',
                    'L.Traffic.DL.PktUuLoss.Tot.QCI.1 (packet)': 'int64',
                    'L.Traffic.UL.PktLoss.Loss (packet)': 'int64',
                    'L.Traffic.UL.PktLoss.Tot (packet)': 'int64',
                    'L.Traffic.UL.PktLoss.Loss.QCI.1 (packet)': 'int64',
                    'L.Traffic.UL.PktLoss.Tot.QCI.1 (packet)': 'int64',
                    'L.E-RAB.SessionTime.HighPrecision.QCI1 (100 ms)': 'int64'           
                })
                self.lteCounters1_df = self.lteCounters1_df.drop_duplicates(subset=['Start Time', 'Cell'])
                print('LTE Queries 1 imported successfully')
            # Query 2
            excel_files = glob.glob(os.path.join(self.lte_path_2, "*.csv"))
            appended_data = []
            print('Importing LTE Queries (2)...')
            for f in excel_files:
                df = pd.read_csv(f, skiprows=7, converters={
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'L.UL.Interference.Avg (dBm)': conv_interference,
                    'L.UL.Interference.Max (dBm)': conv_interference,
                    'L.UL.Interference.Min (dBm)': conv_interference,
                    'L.ChMeas.CQI.DL.0 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.1 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.10 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.11 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.12 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.13 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.14 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.15 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.2 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.3 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.4 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.5 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.6 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.7 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.8 (None)': conv_number_str,
                    'L.ChMeas.CQI.DL.9 (None)': conv_number_str
                })
                df.drop(['Period(min)'], axis=1, inplace=True)
                appended_data.append(df)
                print('Imported LTE query 2: '+f)
            if appended_data:
                self.lteCounters2_df = pd.concat(appended_data)
                self.lteCounters2_df = self.lteCounters2_df.astype({
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'L.UL.Interference.Avg (dBm)': 'float64',
                    'L.UL.Interference.Max (dBm)': 'float64',
                    'L.UL.Interference.Min (dBm)': 'float64',
                    'L.ChMeas.CQI.DL.0 (None)': 'int64',
                    'L.ChMeas.CQI.DL.1 (None)': 'int64',
                    'L.ChMeas.CQI.DL.10 (None)': 'int64',
                    'L.ChMeas.CQI.DL.11 (None)': 'int64',
                    'L.ChMeas.CQI.DL.12 (None)': 'int64',
                    'L.ChMeas.CQI.DL.13 (None)': 'int64',
                    'L.ChMeas.CQI.DL.14 (None)': 'int64',
                    'L.ChMeas.CQI.DL.15 (None)': 'int64',
                    'L.ChMeas.CQI.DL.2 (None)': 'int64',
                    'L.ChMeas.CQI.DL.3 (None)': 'int64',
                    'L.ChMeas.CQI.DL.4 (None)': 'int64',
                    'L.ChMeas.CQI.DL.5 (None)': 'int64',
                    'L.ChMeas.CQI.DL.6 (None)': 'int64',
                    'L.ChMeas.CQI.DL.7 (None)': 'int64',
                    'L.ChMeas.CQI.DL.8 (None)': 'int64',
                    'L.ChMeas.CQI.DL.9 (None)': 'int64'
                })
                self.lteCounters2_df = self.lteCounters2_df.drop_duplicates(subset=['Start Time', 'Cell'])
                print('LTE Queries 2 imported successfully')
            # Query MOCN
            excel_files = glob.glob(os.path.join(self.lte_path_mocn, "*.csv"))
            appended_data = []
            print('Importing LTE Queries (MOCN)...')
            for f in excel_files:
                df = pd.read_csv(f, skiprows=7, converters={
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'CnOperator': str,
                    'L.Traffic.User.Avg.PLMN (None)': conv_number_str,
                    'L.Thrp.bits.DL.PLMN (bit)': conv_number_str,
                    'L.Thrp.bits.UL.PLMN (bit)': conv_number_str,
                    'L.E-RAB.SessionTime.HighPrecision.PLMN.QCI1 (100 ms)': conv_number_str,
                    'L.Thrp.bits.DL.PLMN.QCI.8 (bit)': conv_number_str,
                    'L.Thrp.bits.UL.PLMN.QCI.8 (bit)': conv_number_str,
                    'L.E-RAB.SuccEst.PLMN (None)': conv_number_str,
                    'L.E-RAB.AttEst.PLMN (None)': conv_number_str,
                    'L.E-RAB.SuccEst.PLMN.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.AttEst.PLMN.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.AbnormRel.PLMN (None)': conv_number_str,
                    'L.E-RAB.NormRel.PLMN (None)': conv_number_str,
                    'L.IRATHO.E2W.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.E-RAB.AbnormRel.MME.PLMN (None)': conv_number_str,
                    'L.E-RAB.AbnormRel.PLMN.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.NormRel.PLMN.QCI.1 (None)': conv_number_str,
                    'L.E-RAB.AbnormRel.MME.VoIP.PLMN (None)': conv_number_str,
                    'L.Thrp.bits.DL.LastTTI.PLMN (bit)': conv_number_str,
                    'L.Thrp.Time.DL.RmvLastTTI.PLMN (ms)': conv_number_str,
                    'L.Thrp.bits.UE.UL.LastTTI.PLMN (bit)': conv_number_str,
                    'L.Thrp.Time.UE.UL.RmvLastTTI.PLMN (ms)': conv_number_str,
                    'L.RBUsedOwn.DL.PLMN (None)': conv_number_str,
                    'L.RBUsedOwn.UL.PLMN (None)': conv_number_str,
                    'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntraeNB.IntraFreq.PrepAttOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntraeNB.InterFreq.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntraeNB.InterFreq.PrepAttOut.PLMN (None)': conv_number_str,
                    'L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)': conv_number_str,
                    'L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)': conv_number_str,
                    'L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)': conv_number_str,
                    'L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)': conv_number_str,
                    'L.CSFB.PrepSucc.PLMN (None)': conv_number_str,
                    'L.CSFB.PrepAtt.PLMN (None)': conv_number_str
                })
                df.drop(['Period(min)'], axis=1, inplace=True)
                appended_data.append(df)
                print('Imported LTE query MOCN: '+f)
            if appended_data:
                self.lteCountersMocn_df = pd.concat(appended_data)
                self.lteCountersMocn_df = self.lteCountersMocn_df.astype({
                    'Start Time': str,
                    'NE Name': str,
                    'Cell': str,
                    'CnOperator': str,
                    'L.Traffic.User.Avg.PLMN (None)': 'float64',
                    'L.Thrp.bits.DL.PLMN (bit)': 'float64',
                    'L.Thrp.bits.UL.PLMN (bit)': 'float64',
                    'L.E-RAB.SessionTime.HighPrecision.PLMN.QCI1 (100 ms)': 'int64',
                    'L.Thrp.bits.DL.PLMN.QCI.8 (bit)': 'float64',
                    'L.Thrp.bits.UL.PLMN.QCI.8 (bit)': 'float64',
                    'L.E-RAB.SuccEst.PLMN (None)': 'int64',
                    'L.E-RAB.AttEst.PLMN (None)': 'int64',
                    'L.E-RAB.SuccEst.PLMN.QCI.1 (None)': 'int64',
                    'L.E-RAB.AttEst.PLMN.QCI.1 (None)': 'int64',
                    'L.E-RAB.AbnormRel.PLMN (None)': 'int64',
                    'L.E-RAB.NormRel.PLMN (None)': 'int64',
                    'L.IRATHO.E2W.ExecSuccOut.PLMN (None)': 'int64',
                    'L.E-RAB.AbnormRel.MME.PLMN (None)': 'int64',
                    'L.E-RAB.AbnormRel.PLMN.QCI.1 (None)': 'int64',
                    'L.E-RAB.NormRel.PLMN.QCI.1 (None)': 'int64',
                    'L.E-RAB.AbnormRel.MME.VoIP.PLMN (None)': 'int64',
                    'L.Thrp.bits.DL.LastTTI.PLMN (bit)': 'float64',
                    'L.Thrp.Time.DL.RmvLastTTI.PLMN (ms)': 'int64',
                    'L.Thrp.bits.UE.UL.LastTTI.PLMN (bit)': 'float64',
                    'L.Thrp.Time.UE.UL.RmvLastTTI.PLMN (ms)': 'int64',
                    'L.RBUsedOwn.DL.PLMN (None)': 'float64',
                    'L.RBUsedOwn.UL.PLMN (None)': 'float64',
                    'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.PLMN (None)': 'int64',
                    'L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)': 'int64',
                    'L.HHO.IntraeNB.IntraFreq.PrepAttOut.PLMN (None)': 'int64',
                    'L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)': 'int64',
                    'L.HHO.IntraeNB.InterFreq.ExecSuccOut.PLMN (None)': 'int64',
                    'L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)': 'int64',
                    'L.HHO.IntraeNB.InterFreq.PrepAttOut.PLMN (None)': 'int64',
                    'L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)': 'int64',
                    'L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)': 'int64',
                    'L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)': 'int64',
                    'L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)': 'int64',
                    'L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)': 'int64',
                    'L.CSFB.PrepSucc.PLMN (None)': 'int64',
                    'L.CSFB.PrepAtt.PLMN (None)': 'int64'
                })
                self.lteCountersMocn_df = self.lteCountersMocn_df.drop_duplicates(subset=['Start Time', 'Cell', 'CnOperator'])
                print('LTE Queries MOCN imported successfully')
            # lte_kpi = LTEReportGenerator(self.lteCounters1_df, self.lteCounters2_df, self.lteCountersMocn_df)
            # lte_kpi.generate_excel_report()
            lte_kpi = LTEReportGenerator()
            lte_kpi.generate_report_data(self.lteCounters1_df, self.lteCounters2_df, self.lteCountersMocn_df)
