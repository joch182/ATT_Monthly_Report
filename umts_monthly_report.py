import pandas as pd
from excel_report_writer import ReportWriter
from helpers import count_per_NE

def get_cell_day_kpi(df_concatenated_grouped):
    
    df_concatenated_grouped['VS.RRC.Setup.Succ.Ration.Server.CELL.custom'] = round(100*(df_concatenated_grouped['RRC.SuccConnEstab.OrgConvCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgStrCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgInterCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgBkgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgSubCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmConvCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmStrCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmItrCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmBkgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.EmgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.Unkown (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgHhPrSig (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgLwPrSig (None)']+df_concatenated_grouped['RRC.SuccConnEstab.CallReEst (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmHhPrSig (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmLwPrSig (None)'])/(df_concatenated_grouped['RRC.AttConnEstab.OrgConvCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgStrCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgInterCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgBkgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgSubCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmConvCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmStrCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmInterCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmBkgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.EmgCall (None)']+ df_concatenated_grouped['RRC.AttConnEstab.Unknown (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgHhPrSig (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgLwPrSig (None)']+df_concatenated_grouped['RRC.AttConnEstab.CallReEst (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmHhPrSig (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmLwPrSig (None)']),2)

    df_concatenated_grouped['Service.VS.RRC.Setup.Succ.Ration.custom'] = round(100*((df_concatenated_grouped['RRC.SuccConnEstab.OrgConvCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgStrCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgInterCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgBkgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgSubCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmConvCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmStrCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmItrCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmBkgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.EmgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.Unkown (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgHhPrSig (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgLwPrSig (None)']+df_concatenated_grouped['RRC.SuccConnEstab.CallReEst (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmHhPrSig (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmLwPrSig (None)'])+(df_concatenated_grouped['VS.SuccCellUpdt.PageRsp (None)']+df_concatenated_grouped['VS.SuccCellUpdt.ULDataTrans (None)']+df_concatenated_grouped['VS.SuccCellUpdt.Reg.PCH (None)']+df_concatenated_grouped['VS.SuccCellUpdt.Detach.PCH (None)']+df_concatenated_grouped['VS.SuccCellUpdt.Other.PCH (None)']))/((df_concatenated_grouped['RRC.AttConnEstab.OrgConvCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgStrCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgInterCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgBkgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgSubCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmConvCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmStrCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmInterCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmBkgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.EmgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.Unknown (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgHhPrSig (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgLwPrSig (None)']+df_concatenated_grouped['RRC.AttConnEstab.CallReEst (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmHhPrSig (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmLwPrSig (None)'])+(df_concatenated_grouped['VS.AttCellUpdt.PageRsp (None)']+df_concatenated_grouped['VS.AttCellUpdt.ULDataTrans (None)']+df_concatenated_grouped['VS.AttCellUpdt.Reg.PCH (None)']+df_concatenated_grouped['VS.AttCellUpdt.Detach.PCH (None)']+df_concatenated_grouped['VS.AttCellUpdt.Other.PCH (None)'])),2)

    df_concatenated_grouped['VS.RRC.Congestion.Ratio.custom'] = round(100*(df_concatenated_grouped['VS.RRC.Rej.ULPower.Cong (None)']+df_concatenated_grouped['VS.RRC.Rej.DLPower.Cong (None)']+df_concatenated_grouped['VS.RRC.Rej.ULIUBBand.Cong (None)']+df_concatenated_grouped['VS.RRC.Rej.DLIUBBand.Cong (None)']+df_concatenated_grouped['VS.RRC.Rej.ULCE.Cong (None)']+df_concatenated_grouped['VS.RRC.Rej.DLCE.Cong (None)']+df_concatenated_grouped['VS.RRC.Rej.Code.Cong (None)'])/(df_concatenated_grouped['VS.RRC.AttConnEstab.Sum (None)']),2)

    df_concatenated_grouped['VS.CS.RAB.Setup.Succ.Ration.CELL.custom'] = round(100*(df_concatenated_grouped['VS.RAB.SuccEstabCS.Conv (None)']+df_concatenated_grouped['VS.RAB.SuccEstabCS.Str (None)'])/(df_concatenated_grouped['VS.RAB.AttEstabCS.Conv (None)']+df_concatenated_grouped['VS.RAB.AttEstabCS.Str (None)']),2)

    df_concatenated_grouped['VS.PS.RAB.Setup.Succ.Ration.CELL.custom'] = round(100*(df_concatenated_grouped['VS.RAB.SuccEstabPS.Conv (None)']+df_concatenated_grouped['VS.RAB.SuccEstabPS.Str (None)']+df_concatenated_grouped['VS.RAB.SuccEstabPS.Int (None)']+df_concatenated_grouped['VS.RAB.SuccEstabPS.Bkg (None)'])/(df_concatenated_grouped['VS.RAB.AttEstabPS.Conv (None)']+df_concatenated_grouped['VS.RAB.AttEstabPS.Str (None)']+df_concatenated_grouped['VS.RAB.AttEstabPS.Int (None)']+df_concatenated_grouped['VS.RAB.AttEstabPS.Bkg (None)']),2)

    df_concatenated_grouped['VS.CS.Radio.Access.Succ.Ratio.CELL.custom'] = round(100*(df_concatenated_grouped['RRC.SuccConnEstab.OrgConvCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmConvCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.EmgCall (None)'])/(df_concatenated_grouped['RRC.AttConnEstab.OrgConvCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmConvCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.EmgCall (None)'])*(df_concatenated_grouped['VS.RAB.SuccEstabCS.Conv (None)']+df_concatenated_grouped['VS.RAB.SuccEstabCS.Str (None)'])/(df_concatenated_grouped['VS.RAB.AttEstabCS.Conv (None)']+df_concatenated_grouped['VS.RAB.AttEstabCS.Str (None)']),2)

    df_concatenated_grouped['VS.PS.Radio.Access.Succ.Ratio.CELL.custom'] = round(100*(df_concatenated_grouped['RRC.SuccConnEstab.OrgBkgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.OrgInterCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmBkgCall (None)']+df_concatenated_grouped['RRC.SuccConnEstab.TmItrCall (None)'])/(df_concatenated_grouped['RRC.AttConnEstab.OrgBkgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.OrgInterCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmBkgCall (None)']+df_concatenated_grouped['RRC.AttConnEstab.TmInterCall (None)'])*(df_concatenated_grouped['VS.RAB.SuccEstabPS.Conv (None)']+df_concatenated_grouped['VS.RAB.SuccEstabPS.Str (None)']+df_concatenated_grouped['VS.RAB.SuccEstabPS.Int (None)']+df_concatenated_grouped['VS.RAB.SuccEstabPS.Bkg (None)'])/(df_concatenated_grouped['VS.RAB.AttEstabPS.Conv (None)']+df_concatenated_grouped['VS.RAB.AttEstabPS.Str (None)']+df_concatenated_grouped['VS.RAB.AttEstabPS.Int (None)']+df_concatenated_grouped['VS.RAB.AttEstabPS.Bkg (None)']),2)

    df_concatenated_grouped['IU.Paging.Congestion.Cell.Ratio'] = round(100*(df_concatenated_grouped['VS.RRC.Paging1.Loss.PCHCong.Cell (None)']/df_concatenated_grouped['VS.UTRAN.AttPaging1 (None)']),2)

    df_concatenated_grouped['VS.CS.Call.Drop.Cell.Rate.custom'] = round(100*(df_concatenated_grouped['VS.RAB.AbnormRel.CS (None)']/(df_concatenated_grouped['VS.RAB.AbnormRel.CS (None)']+df_concatenated_grouped['VS.RAB.NormRel.CS (None)'])),2)

    df_concatenated_grouped['VS.PS.Call.Drop.Cell.Rate.custom'] = round(100*(df_concatenated_grouped['VS.RAB.AbnormRel.PS (None)']/(df_concatenated_grouped['VS.RAB.AbnormRel.PS (None)']+df_concatenated_grouped['VS.RAB.NormRel.PS (None)'])),2)

    df_concatenated_grouped['VS.HSDPA.RAB.ABNORMREL.RATE.A.custom'] = round(100*(df_concatenated_grouped['VS.HSDPA.RAB.AbnormRel (None)']-df_concatenated_grouped['VS.HSDPA.RAB.AbnormRel.H2P (None)'])/(df_concatenated_grouped['VS.HSDPA.RAB.AbnormRel (None)']+df_concatenated_grouped['VS.HSDPA.RAB.NormRel (None)']+df_concatenated_grouped['VS.HSDPA.HHO.H2D.SuccOutIntraFreq (None)']+df_concatenated_grouped['VS.HSDPA.HHO.H2D.SuccOutInterFreq (None)']+df_concatenated_grouped['VS.HSDPA.H2D.Succ (None)']+df_concatenated_grouped['VS.HSDPA.H2F.Succ (None)']+df_concatenated_grouped['VS.HSDPA.H2P.Succ (None)']),2)

    df_concatenated_grouped['VS.HSUPA.RAB.ABNORMREL.RATE.A.custom'] = round(100*(df_concatenated_grouped['VS.HSUPA.RAB.AbnormRel (None)']-df_concatenated_grouped['VS.HSUPA.RAB.AbnormRel.E2P (None)'])/(df_concatenated_grouped['VS.HSUPA.RAB.AbnormRel (None)']+df_concatenated_grouped['VS.HSUPA.RAB.NormRel (None)']+df_concatenated_grouped['VS.HSUPA.HHO.E2D.SuccOutIntraFreq (None)']+df_concatenated_grouped['VS.HSUPA.HHO.E2D.SuccOutInterFreq (None)']+df_concatenated_grouped['VS.HSUPA.E2F.Succ (None)']+df_concatenated_grouped['VS.HSUPA.E2D.Succ (None)']+df_concatenated_grouped['VS.HSUPA.E2P.Succ (None)']),2)

    df_concatenated_grouped['VS.R99.RAB.ABNORMREL.RATE.custom'] = round(100*df_concatenated_grouped['VS.RAB.AbnormRel.PSR99 (None)']/(df_concatenated_grouped['VS.RAB.AbnormRel.PSR99 (None)']+df_concatenated_grouped['VS.RAB.NormRel.PSR99 (None)']),2)

    df_concatenated_grouped['VS.SHO.Success.Cell.Rate.custom'] = round(100*(df_concatenated_grouped['VS.SHO.SuccRLAdd (None)']+df_concatenated_grouped['VS.SHO.SuccRLDel (None)'])/(df_concatenated_grouped['VS.SHO.AttRLAdd (None)']+df_concatenated_grouped['VS.SHO.AttRLDel (None)']),2)

    df_concatenated_grouped['VS.HHO.IntraFreqOut.Succ.Cell.Rate.custom'] = round(100*(df_concatenated_grouped['VS.HHO.SuccIntraFreqOut.IntraNodeB (None)']+df_concatenated_grouped['VS.HHO.SuccIntraFreqOut.InterNodeBIntraRNC (None)']+df_concatenated_grouped['VS.HHO.SuccIntraFreqOut.InterRNC (None)'])/(df_concatenated_grouped['VS.HHO.AttIntraFreqOut.InterNodeBIntraRNC (None)']+df_concatenated_grouped['VS.HHO.AttIntraFreqOut.InterRNC (None)']+df_concatenated_grouped['VS.HHO.AttIntraFreqOut.IntraNodeB (None)']),2)

    df_concatenated_grouped['VS.HHO.InterFreqOut.Succ.Cell.Rate.custom'] = round(100*(df_concatenated_grouped['VS.HHO.SuccInterFreqOut (None)']/df_concatenated_grouped['VS.HHO.AttInterFreqOut (None)']),2)

    df_concatenated_grouped['R99.CODE.Utilization.custom'] = round(100*(((df_concatenated_grouped['VS.SingleRAB.SF4 (None)']+df_concatenated_grouped['VS.MultRAB.SF4 (None)'])*64)+((df_concatenated_grouped['VS.SingleRAB.SF8 (None)']+df_concatenated_grouped['VS.MultRAB.SF8 (None)'])*32)+((df_concatenated_grouped['VS.SingleRAB.SF16 (None)']+df_concatenated_grouped['VS.MultRAB.SF16 (None)'])*16)+((df_concatenated_grouped['VS.SingleRAB.SF32 (None)']+df_concatenated_grouped['VS.MultRAB.SF32 (None)'])*8)+((df_concatenated_grouped['VS.SingleRAB.SF64 (None)']+df_concatenated_grouped['VS.MultRAB.SF64 (None)'])*4)+((df_concatenated_grouped['VS.SingleRAB.SF128 (None)']+df_concatenated_grouped['VS.MultRAB.SF128 (None)'])*2)+((df_concatenated_grouped['VS.SingleRAB.SF256 (None)']+df_concatenated_grouped['VS.MultRAB.SF256 (None)'])))/256,2)

    df_concatenated_grouped['R99.Traffic.custom'] = round(((df_concatenated_grouped['VS.PS.Bkg.DL.8.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.16.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.32.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.64.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.128.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.144.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.256.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.DL.384.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.8.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.16.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.32.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.64.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.128.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.144.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.256.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.DL.384.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.8.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.16.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.32.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.64.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.128.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.144.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.256.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.DL.384.Traffic (bit)']+df_concatenated_grouped['VS.PS.Conv.DL.Traffic (bit)'])/1048576/8)+((df_concatenated_grouped['VS.PS.Bkg.UL.8.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.16.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.32.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.64.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.128.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.144.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.256.Traffic (bit)']+df_concatenated_grouped['VS.PS.Bkg.UL.384.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.8.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.16.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.32.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.64.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.128.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.144.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.256.Traffic (bit)']+df_concatenated_grouped['VS.PS.Int.UL.384.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.UL.8.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.UL.16.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.UL.32.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.UL.64.Traffic (bit)']+df_concatenated_grouped['VS.PS.Str.UL.128.Traffic (bit)']+df_concatenated_grouped['VS.PS.Conv.UL.Traffic (bit)'])/1048576/8),2)

    df_concatenated_grouped['Traffic.HS.custom'] = round((df_concatenated_grouped['VS.HSDPA.MeanChThroughput.TotalBytes (byte)']+df_concatenated_grouped['VS.HSUPA.MeanChThroughput.TotalBytes (byte)'])/1048576,2)

    # Final data to return
    return df_concatenated_grouped[['DATE', 'VS.RRC.Setup.Succ.Ration.Server.CELL.custom', 'Service.VS.RRC.Setup.Succ.Ration.custom', 'VS.RRC.Congestion.Ratio.custom','VS.CS.RAB.Setup.Succ.Ration.CELL.custom','VS.PS.RAB.Setup.Succ.Ration.CELL.custom','VS.CS.Radio.Access.Succ.Ratio.CELL.custom','VS.PS.Radio.Access.Succ.Ratio.CELL.custom','IU.Paging.Congestion.Cell.Ratio','VS.CS.Call.Drop.Cell.Rate.custom','VS.PS.Call.Drop.Cell.Rate.custom','VS.HSDPA.RAB.ABNORMREL.RATE.A.custom','VS.HSUPA.RAB.ABNORMREL.RATE.A.custom','VS.R99.RAB.ABNORMREL.RATE.custom','VS.SHO.Success.Cell.Rate.custom','VS.HHO.IntraFreqOut.Succ.Cell.Rate.custom','VS.HHO.InterFreqOut.Succ.Cell.Rate.custom','VS.HSDPA.MeanChThroughput (kbit/s)','VS.HSUPA.MeanChThroughput (kbit/s)','VS.Cell.UnavailTime (s)','VS.Cell.UnavailTime.Sys (s)','VS.MeanRTWP (dBm)','VS.MaxRTWP (dBm)','VS.MinRTWP (dBm)','VS.MeanTCP (dBm)','VS.MeanTCP.NonHS (dBm)','R99.CODE.Utilization.custom','VS.RAB.AMR.Erlang.cell (Erl)','VS.RAB.AMRWB.Erlang.cell (Erl)','R99.Traffic.custom','Traffic.HS.custom','VS.CellDCHUEs (None)','VS.HSDPA.UE.Mean.Cell (None)','VS.HSUPA.UE.Mean.Cell (None)','VS.PSLoad.ULThruput.MOCN.PLMN0 (byte)','VS.PSLoad.ULThruput.MOCN.PLMN1 (byte)','VS.PSLoad.ULThruput.MOCN.PLMN2 (byte)','VS.PSLoad.ULThruput.MOCN.PLMN3 (byte)','VS.PSLoad.DLThruput.MOCN.PLMN0 (byte)','VS.PSLoad.DLThruput.MOCN.PLMN1 (byte)','VS.PSLoad.DLThruput.MOCN.PLMN2 (byte)','VS.PSLoad.DLThruput.MOCN.PLMN3 (byte)','VS.CS.Erlang.Equiv.MOCN.PLMN0 (Erl)','VS.CS.Erlang.Equiv.MOCN.PLMN1 (Erl)','VS.CS.Erlang.Equiv.MOCN.PLMN2 (Erl)','VS.CS.Erlang.Equiv.MOCN.PLMN3 (Erl)']]

def get_cell_bh_kpi(df_concatenated_grouped):
    
    df_concatenated_grouped['R99.CODE.Utilization.custom'] = round(100*(((df_concatenated_grouped['VS.SingleRAB.SF4 (None)']+df_concatenated_grouped['VS.MultRAB.SF4 (None)'])*64)+((df_concatenated_grouped['VS.SingleRAB.SF8 (None)']+df_concatenated_grouped['VS.MultRAB.SF8 (None)'])*32)+((df_concatenated_grouped['VS.SingleRAB.SF16 (None)']+df_concatenated_grouped['VS.MultRAB.SF16 (None)'])*16)+((df_concatenated_grouped['VS.SingleRAB.SF32 (None)']+df_concatenated_grouped['VS.MultRAB.SF32 (None)'])*8)+((df_concatenated_grouped['VS.SingleRAB.SF64 (None)']+df_concatenated_grouped['VS.MultRAB.SF64 (None)'])*4)+((df_concatenated_grouped['VS.SingleRAB.SF128 (None)']+df_concatenated_grouped['VS.MultRAB.SF128 (None)'])*2)+((df_concatenated_grouped['VS.SingleRAB.SF256 (None)']+df_concatenated_grouped['VS.MultRAB.SF256 (None)'])))/256,2)

    # Final data to return
    return df_concatenated_grouped[['DATE', 'VS.MeanTCP (dBm)','VS.MeanTCP.NonHS (dBm)','R99.CODE.Utilization.custom']]

def merge_df(df1, df2):
    return pd.concat([df1, df2], axis=1, join="inner").reset_index()

def get_BH_df(df):
    # Sort PRB usage descending and then remove duplicates for each date/cell
    # Cell data
    df = df.reset_index()
    df = df.sort_values('VS.MeanTCP (dBm)', ascending=False).drop_duplicates(['DATE', 'CellName'])
    df = df.groupby(by=['DATE']).agg({
        'VS.MeanTCP (dBm)': 'mean',
        'VS.MeanTCP.NonHS (dBm)': 'mean',
        'VS.MultRAB.SF128 (None)':  'mean',
        'VS.MultRAB.SF16 (None)': 'mean',
        'VS.MultRAB.SF256 (None)': 'mean',
        'VS.MultRAB.SF32 (None)': 'mean',
        'VS.MultRAB.SF4 (None)': 'mean',
        'VS.MultRAB.SF64 (None)': 'mean',
        'VS.MultRAB.SF8 (None)': 'mean',
        'VS.SingleRAB.SF128 (None)': 'mean',
        'VS.SingleRAB.SF16 (None)': 'mean',
        'VS.SingleRAB.SF256 (None)': 'mean',
        'VS.SingleRAB.SF32 (None)': 'mean',
        'VS.SingleRAB.SF4 (None)': 'mean',
        'VS.SingleRAB.SF64 (None)': 'mean',
        'VS.SingleRAB.SF8 (None)': 'mean'
    })
    return df.reset_index()

class UMTSReportCalc():
    def __init__(self, df) -> None:
        self.umtsCellCounters_df = df.copy()
        self.umtsCellCounters_df['Start Time'] = self.umtsCellCounters_df['Start Time'].str.replace(':00 DST', '')
        self.umtsCellCounters_df['Start Time'] = pd.to_datetime(self.umtsCellCounters_df['Start Time'])
        self.umtsCellCounters_df['DATE'] = pd.to_datetime(self.umtsCellCounters_df['Start Time']).dt.date
        self.umtsCellCounters_df['Site'] = self.umtsCellCounters_df['Cell'].str.split(pat="=", expand=True)[1].str.split(pat="_", expand=True)[0]
        self.umtsCellCounters_df['CellName'] = self.umtsCellCounters_df['Cell'].str.split(pat="=", expand=True)[1].str.split(pat=",", expand=True)[0]
        self.umtsCellCounters_df['CI'] = self.umtsCellCounters_df['Cell'].str.split(pat="=", expand=True)[2].str.split(pat=",", expand=True)[0]
        self.umtsCellCounters_df = self.umtsCellCounters_df.drop(['Cell'], axis=1)
        self.umtsCellCounters_df = self.umtsCellCounters_df.astype({"CI": int})

    def group_data(self, query, by='ALL_DAY'):
        if by == 'ALL_DAY':
            group_cell_by = ['DATE']
        elif by == 'BH':
            group_cell_by = ['DATE', 'Start Time', 'CellName']

        if query == 1:
            self.umtsCellCounters_df = self.umtsCellCounters_df.groupby(by=group_cell_by).agg({
                'VS.MultRAB.SF128 (None)':  'mean',
                'VS.MultRAB.SF16 (None)': 'mean',
                'VS.MultRAB.SF256 (None)': 'mean',
                'VS.MultRAB.SF32 (None)': 'mean',
                'VS.MultRAB.SF4 (None)': 'mean',
                'VS.MultRAB.SF64 (None)': 'mean',
                'VS.MultRAB.SF8 (None)': 'mean',
                'VS.SingleRAB.SF128 (None)': 'mean',
                'VS.SingleRAB.SF16 (None)': 'mean',
                'VS.SingleRAB.SF256 (None)': 'mean',
                'VS.SingleRAB.SF32 (None)': 'mean',
                'VS.SingleRAB.SF4 (None)': 'mean',
                'VS.SingleRAB.SF64 (None)': 'mean',
                'VS.SingleRAB.SF8 (None)': 'mean',
                'RRC.SuccConnEstab.OrgConvCall (None)': 'sum',
                'RRC.SuccConnEstab.OrgStrCall (None)': 'sum',
                'RRC.SuccConnEstab.OrgInterCall (None)': 'sum',
                'RRC.SuccConnEstab.OrgBkgCall (None)': 'sum',
                'RRC.SuccConnEstab.OrgSubCall (None)': 'sum',
                'RRC.SuccConnEstab.TmConvCall (None)': 'sum',
                'RRC.SuccConnEstab.TmStrCall (None)': 'sum',
                'RRC.SuccConnEstab.TmItrCall (None)': 'sum',
                'RRC.SuccConnEstab.TmBkgCall (None)': 'sum',
                'RRC.SuccConnEstab.EmgCall (None)': 'sum',
                'RRC.SuccConnEstab.Unkown (None)': 'sum',
                'RRC.SuccConnEstab.OrgHhPrSig (None)': 'sum',
                'RRC.SuccConnEstab.OrgLwPrSig (None)': 'sum',
                'RRC.SuccConnEstab.CallReEst (None)': 'sum',
                'RRC.SuccConnEstab.TmHhPrSig (None)': 'sum',
                'RRC.SuccConnEstab.TmLwPrSig (None)': 'sum',
                'RRC.AttConnEstab.OrgConvCall (None)': 'sum',
                'RRC.AttConnEstab.OrgInterCall (None)': 'sum',
                'RRC.AttConnEstab.OrgStrCall (None)': 'sum',
                'RRC.AttConnEstab.OrgBkgCall (None)': 'sum',
                'RRC.AttConnEstab.OrgSubCall (None)': 'sum',
                'RRC.AttConnEstab.TmBkgCall (None)': 'sum',
                'RRC.AttConnEstab.TmConvCall (None)': 'sum',
                'RRC.AttConnEstab.TmInterCall (None)': 'sum',
                'RRC.AttConnEstab.TmStrCall (None)': 'sum',
                'RRC.AttConnEstab.EmgCall (None)': 'sum',
                'RRC.AttConnEstab.Unknown (None)': 'sum',
                'RRC.AttConnEstab.CallReEst (None)': 'sum',
                'RRC.AttConnEstab.OrgHhPrSig (None)': 'sum',
                'RRC.AttConnEstab.OrgLwPrSig (None)': 'sum',
                'RRC.AttConnEstab.TmHhPrSig (None)': 'sum',
                'RRC.AttConnEstab.TmLwPrSig (None)': 'sum',
                'VS.SuccCellUpdt.PageRsp (None)': 'sum',
                'VS.SuccCellUpdt.ULDataTrans (None)': 'sum',
                'VS.SuccCellUpdt.Reg.PCH (None)': 'sum',
                'VS.SuccCellUpdt.Detach.PCH (None)': 'sum',
                'VS.SuccCellUpdt.Other.PCH (None)': 'sum',
                'VS.AttCellUpdt.PageRsp (None)': 'sum',
                'VS.AttCellUpdt.ULDataTrans (None)': 'sum',
                'VS.AttCellUpdt.Reg.PCH (None)': 'sum',
                'VS.AttCellUpdt.Detach.PCH (None)': 'sum',
                'VS.AttCellUpdt.Other.PCH (None)': 'sum',
                'VS.RRC.Rej.ULPower.Cong (None)': 'sum',
                'VS.RRC.Rej.DLPower.Cong (None)': 'sum',
                'VS.RRC.Rej.ULIUBBand.Cong (None)': 'sum',
                'VS.RRC.Rej.DLIUBBand.Cong (None)': 'sum',
                'VS.RRC.Rej.ULCE.Cong (None)': 'sum',
                'VS.RRC.Rej.DLCE.Cong (None)': 'sum',
                'VS.RRC.Rej.Code.Cong (None)': 'sum',
                'VS.RRC.AttConnEstab.Sum (None)': 'sum',
                'VS.RAB.SuccEstabCS.Conv (None)': 'sum',
                'VS.RAB.SuccEstabCS.Str (None)': 'sum',
                'VS.RAB.AttEstabCS.Conv (None)': 'sum',
                'VS.RAB.AttEstabCS.Str (None)': 'sum',
                'VS.RAB.SuccEstabPS.Conv (None)': 'sum',
                'VS.RAB.SuccEstabPS.Str (None)': 'sum',
                'VS.RAB.SuccEstabPS.Int (None)': 'sum',
                'VS.RAB.SuccEstabPS.Bkg (None)': 'sum',
                'VS.RAB.AttEstabPS.Conv (None)': 'sum',
                'VS.RAB.AttEstabPS.Str (None)': 'sum',
                'VS.RAB.AttEstabPS.Int (None)': 'sum',
                'VS.RAB.AttEstabPS.Bkg (None)': 'sum',
                'VS.RRC.Paging1.Loss.PCHCong.Cell (None)': 'sum',
                'VS.UTRAN.AttPaging1 (None)': 'sum',
                'VS.RAB.AbnormRel.CS (None)': 'sum',
                'VS.RAB.NormRel.CS (None)': 'sum',
                'VS.RAB.AbnormRel.PS (None)': 'sum',
                'VS.RAB.NormRel.PS (None)': 'sum',
                'VS.HSDPA.RAB.AbnormRel (None)': 'sum',
                'VS.HSDPA.RAB.AbnormRel.H2P (None)': 'sum',
                'VS.HSDPA.RAB.NormRel (None)': 'sum',
                'VS.HSDPA.HHO.H2D.SuccOutIntraFreq (None)': 'sum',
                'VS.HSDPA.HHO.H2D.SuccOutInterFreq (None)': 'sum',
                'VS.HSDPA.H2D.Succ (None)': 'sum',
                'VS.HSDPA.H2F.Succ (None)': 'sum',
                'VS.HSDPA.H2P.Succ (None)': 'sum',
                'VS.RAB.AbnormRel.PSR99 (None)': 'sum',
                'VS.RAB.NormRel.PSR99 (None)': 'sum',
                'VS.HSDPA.MeanChThroughput (kbit/s)': 'sum',
                'VS.Cell.UnavailTime (s)': 'sum',
                'VS.Cell.UnavailTime.Sys (s)': 'sum',
                'VS.HSDPA.UE.Mean.Cell (None)': 'sum'
                })
        elif query == 2:
            self.umtsCellCounters_df = self.umtsCellCounters_df.groupby(by=group_cell_by).agg({
                'VS.MeanTCP (dBm)': 'mean',
                'VS.MeanTCP.NonHS (dBm)': 'mean',
                'VS.HSUPA.RAB.AbnormRel (None)': 'sum',
                'VS.HSUPA.RAB.AbnormRel.E2P (None)': 'sum',
                'VS.HSUPA.RAB.NormRel (None)': 'sum',
                'VS.HSUPA.HHO.E2D.SuccOutIntraFreq (None)': 'sum',
                'VS.HSUPA.HHO.E2D.SuccOutInterFreq (None)': 'sum',
                'VS.HSUPA.E2F.Succ (None)': 'sum',
                'VS.HSUPA.E2D.Succ (None)': 'sum',
                'VS.HSUPA.E2P.Succ (None)': 'sum',
                'VS.SHO.SuccRLAdd (None)': 'sum',
                'VS.SHO.SuccRLDel (None)': 'sum',
                'VS.SHO.AttRLAdd (None)': 'sum',
                'VS.SHO.AttRLDel (None)': 'sum',
                'VS.HHO.SuccIntraFreqOut.IntraNodeB (None)': 'sum',
                'VS.HHO.SuccIntraFreqOut.InterNodeBIntraRNC (None)': 'sum',
                'VS.HHO.SuccIntraFreqOut.InterRNC (None)': 'sum',
                'VS.HHO.AttIntraFreqOut.InterNodeBIntraRNC (None)': 'sum',
                'VS.HHO.AttIntraFreqOut.InterRNC (None)': 'sum',
                'VS.HHO.AttIntraFreqOut.IntraNodeB (None)': 'sum',
                'VS.HHO.SuccInterFreqOut (None)': 'sum',
                'VS.HHO.AttInterFreqOut (None)': 'sum',
                'VS.HSUPA.MeanChThroughput (kbit/s)': 'sum',
                'VS.MeanRTWP (dBm)': 'mean',
                'VS.MaxRTWP (dBm)': 'mean',
                'VS.MinRTWP (dBm)': 'mean',
                'VS.RAB.AMR.Erlang.cell (Erl)': 'sum',
                'VS.RAB.AMRWB.Erlang.cell (Erl)': 'sum',
                'VS.PS.Bkg.DL.128.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.144.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.16.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.256.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.32.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.384.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.64.Traffic (bit)': 'sum',
                'VS.PS.Bkg.DL.8.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.128.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.144.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.16.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.256.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.32.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.384.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.64.Traffic (bit)': 'sum',
                'VS.PS.Int.DL.8.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.128.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.144.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.16.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.256.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.32.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.384.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.64.Traffic (bit)': 'sum',
                'VS.PS.Str.DL.8.Traffic (bit)': 'sum',
                'VS.PS.Conv.DL.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.128.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.144.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.16.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.256.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.32.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.384.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.64.Traffic (bit)': 'sum',
                'VS.PS.Bkg.UL.8.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.128.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.144.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.16.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.256.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.32.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.384.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.64.Traffic (bit)': 'sum',
                'VS.PS.Int.UL.8.Traffic (bit)': 'sum',
                'VS.PS.Str.UL.128.Traffic (bit)': 'sum',
                'VS.PS.Str.UL.16.Traffic (bit)': 'sum',
                'VS.PS.Str.UL.32.Traffic (bit)': 'sum',
                'VS.PS.Str.UL.64.Traffic (bit)': 'sum',
                'VS.PS.Str.UL.8.Traffic (bit)': 'sum',
                'VS.PS.Conv.UL.Traffic (bit)': 'sum',
                'VS.HSDPA.MeanChThroughput.TotalBytes (byte)': 'sum',
                'VS.HSUPA.MeanChThroughput.TotalBytes (byte)': 'sum',
                'VS.CellDCHUEs (None)': 'sum',
                'VS.HSUPA.UE.Mean.Cell (None)': 'sum',
                'VS.PSLoad.ULThruput.MOCN.PLMN0 (byte)': 'sum',
                'VS.PSLoad.ULThruput.MOCN.PLMN1 (byte)': 'sum',
                'VS.PSLoad.ULThruput.MOCN.PLMN2 (byte)': 'sum',
                'VS.PSLoad.ULThruput.MOCN.PLMN3 (byte)': 'sum',
                'VS.PSLoad.DLThruput.MOCN.PLMN0 (byte)': 'sum',
                'VS.PSLoad.DLThruput.MOCN.PLMN1 (byte)': 'sum',
                'VS.PSLoad.DLThruput.MOCN.PLMN2 (byte)': 'sum',
                'VS.PSLoad.DLThruput.MOCN.PLMN3 (byte)': 'sum',
                'VS.CS.Erlang.Equiv.MOCN.PLMN0 (Erl)': 'sum',
                'VS.CS.Erlang.Equiv.MOCN.PLMN1 (Erl)': 'sum',
                'VS.CS.Erlang.Equiv.MOCN.PLMN2 (Erl)': 'sum',
                'VS.CS.Erlang.Equiv.MOCN.PLMN3 (Erl)': 'sum' 
            })  

class UMTSReportGenerator():

    @staticmethod
    def generate_report_data(df_1, df_2):
        report_writer = ReportWriter()
        df1 = UMTSReportCalc(df_1)
        NE_check_df1 = count_per_NE(df1.umtsCellCounters_df, 'Site')
        Cell_check_df1 = count_per_NE(df1.umtsCellCounters_df, 'CellName')
        # NE Check for Cluster for Query 1
        ws = report_writer.wb.active
        ws.title = 'Data Check NE Q1'
        ws.column_dimensions['A'].width = 19
        report_writer.save_df_sheet(ws, NE_check_df1)
        # Cell Check for Cluster for Query 1
        report_writer.create_sheet('Data Check Cell Q1')
        ws = report_writer.get_sheet('Data Check Cell Q1')
        ws.column_dimensions['A'].width = 19
        report_writer.save_df_sheet(ws, Cell_check_df1)
        del NE_check_df1, Cell_check_df1
        df2 = UMTSReportCalc(df_2)
        NE_check_df2 = count_per_NE(df2.umtsCellCounters_df, 'Site')
        Cell_check_df2 = count_per_NE(df2.umtsCellCounters_df, 'CellName')
        # NE Check for Cluster for Query 2
        report_writer.create_sheet('Data Check NE Q2')
        ws = report_writer.get_sheet('Data Check NE Q2')
        ws.column_dimensions['A'].width = 19
        report_writer.save_df_sheet(ws, NE_check_df2)
        # Cell Check for Cluster for Query 2
        report_writer.create_sheet('Data Check Cell Q2')
        ws = report_writer.get_sheet('Data Check Cell Q2')
        ws.column_dimensions['A'].width = 19
        report_writer.save_df_sheet(ws, Cell_check_df2)
        del NE_check_df2, Cell_check_df2
        df1.group_data(query=1)
        df2.group_data(query=2)
        df_merged = merge_df(df1.umtsCellCounters_df, df2.umtsCellCounters_df)
        del df1, df2
        print('Analysis started...')
         # All day for cluster 
        report_writer.create_sheet('UMTS All Day - Cluster')
        ws = report_writer.get_sheet('UMTS All Day - Cluster')
        kpi_df = get_cell_day_kpi(df_merged)
        report_writer.save_df_sheet(ws, kpi_df)
        print('All day for cluster finished...')
        # BH data for cluster
        df1 = UMTSReportCalc(df_1)
        df1.group_data(query=1, by="BH")
        df2 = UMTSReportCalc(df_2)
        df2.group_data(query=2, by="BH")
        df_merged = merge_df(df1.umtsCellCounters_df, df2.umtsCellCounters_df)
        del df1, df2
        df_merged = get_BH_df(df_merged)
        report_writer.create_sheet('UMTS BH - Cluster')
        ws = report_writer.get_sheet('UMTS BH - Cluster')
        kpi_df = get_cell_bh_kpi(df_merged)
        report_writer.save_df_sheet(ws, kpi_df)
        print('BH for cluster finished...')
        report_writer.save_excel_report("UMTS")
        print('FINISHED')
        