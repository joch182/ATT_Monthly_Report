import pandas as pd
from excel_report_writer import ReportWriter
from helpers import count_per_NE

def get_cell_day_kpi(df_concatenated_grouped):
    df_concatenated_grouped['SIGNALING.L.RRC.SETUP.SUCCESS.RATE.custom'] = round(100*(df_concatenated_grouped['L.RRC.ConnReq.Succ.MoSig (None)']/df_concatenated_grouped['L.RRC.ConnReq.Att.MoSig (None)']),2)

    df_concatenated_grouped['SERVICE.L.RRC.SETUP.SUCCESS.RATE.custom'] =  round(100*(df_concatenated_grouped['L.RRC.ConnReq.Succ.Emc (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.HighPri (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.Mt (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.MoData (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.DelayTol (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.MoVoiceCall (None)'])/(df_concatenated_grouped['L.RRC.ConnReq.Att.Emc (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.HighPri (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.Mt (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.MoData (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.DelayTol (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.MoVoiceCall (None)']),2)

    df_concatenated_grouped['SERVICE.L.RRC.SETUP.SUCCESS.RATE.custom'] = round(100*df_concatenated_grouped['L.E-RAB.SuccEst (None)']/(df_concatenated_grouped['L.E-RAB.AttEst (None)']-df_concatenated_grouped['L.E-RAB.FailEst.X2AP (None)']),2)

    df_concatenated_grouped['ERAB.ESTABLISH.SUCCESS.RATE.custom'] = round(100*df_concatenated_grouped['L.E-RAB.SuccEst (None)']/(df_concatenated_grouped['L.E-RAB.AttEst (None)']-df_concatenated_grouped['L.E-RAB.FailEst.X2AP (None)']),2)

    df_concatenated_grouped['VOIP.ERAB.ESTABLISH.SUCCESS.RATE.custom'] = round(100*df_concatenated_grouped['L.E-RAB.SuccEst.QCI.1 (None)']/(df_concatenated_grouped['L.E-RAB.AttEst.QCI.1 (None)']-df_concatenated_grouped['L.E-RAB.FailEst.X2AP.VoIP (None)']),2)

    df_concatenated_grouped['S1.SIG.CONN.SETUP.SUCCESS.RATE.custom'] = round(100*df_concatenated_grouped['L.S1Sig.ConnEst.Succ (None)']/df_concatenated_grouped['L.S1Sig.ConnEst.Att (None)'],2)

    df_concatenated_grouped['CALL.SETUP.SUCCESS.RATE.custom'] = round(100*((df_concatenated_grouped['L.RRC.ConnReq.Succ.Emc (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.HighPri (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.Mt (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.MoData (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.DelayTol (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.MoVoiceCall (None)'])/(df_concatenated_grouped['L.RRC.ConnReq.Att.Emc (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.HighPri (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.Mt (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.MoData (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.DelayTol (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.MoVoiceCall (None)']))*(df_concatenated_grouped['L.S1Sig.ConnEst.Succ (None)']/df_concatenated_grouped['L.S1Sig.ConnEst.Att (None)'])*(df_concatenated_grouped['L.E-RAB.SuccEst (None)']/(df_concatenated_grouped['L.E-RAB.AttEst (None)']-df_concatenated_grouped['L.E-RAB.FailEst.X2AP (None)'])),2)

    df_concatenated_grouped['VOIP.CALL.SETUP.SUCCESS.RATE.custom'] = round(100*((df_concatenated_grouped['L.RRC.ConnReq.Succ.Emc (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.HighPri (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.Mt (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.MoData (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.DelayTol (None)']+df_concatenated_grouped['L.RRC.ConnReq.Succ.MoVoiceCall (None)'])/(df_concatenated_grouped['L.RRC.ConnReq.Att.Emc (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.HighPri (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.Mt (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.MoData (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.DelayTol (None)']+df_concatenated_grouped['L.RRC.ConnReq.Att.MoVoiceCall (None)']))*(df_concatenated_grouped['L.S1Sig.ConnEst.Succ (None)']/df_concatenated_grouped['L.S1Sig.ConnEst.Att (None)'])*(df_concatenated_grouped['L.E-RAB.SuccEst.QCI.1 (None)']/(df_concatenated_grouped['L.E-RAB.AttEst.QCI.1 (None)']-df_concatenated_grouped['L.E-RAB.FailEst.X2AP (None)'])),2)

    df_concatenated_grouped['SERVICE.RETAINABILITY.custom'] = round(100*(1-(df_concatenated_grouped['L.E-RAB.AbnormRel (None)']/(df_concatenated_grouped['L.E-RAB.AbnormRel (None)']+df_concatenated_grouped['L.E-RAB.NormRel (None)']+df_concatenated_grouped['L.E-RAB.NormRel.IRatHOOut (None)']))),2)

    df_concatenated_grouped['VOIP.CALL.RETAINABILITY.custom'] = round(100*(1-(df_concatenated_grouped['L.E-RAB.AbnormRel.QCI.1 (None)']/(df_concatenated_grouped['L.E-RAB.AbnormRel.QCI.1 (None)']+df_concatenated_grouped['L.E-RAB.NormRel.QCI.1 (None)']+df_concatenated_grouped['L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)']))),2)

    df_concatenated_grouped['INTRA.FREQ.HANDOVER.OUT.SUCCESS.RATE.custom'] = round(100*(df_concatenated_grouped['L.HHO.IntraeNB.IntraFreq.ExecSuccOut (None)']+df_concatenated_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)'])/(df_concatenated_grouped['L.HHO.IntraeNB.IntraFreq.ExecAttOut (None)']+df_concatenated_grouped['L.HHO.IntereNB.IntraFreq.ExecAttOut (None)']),2)

    df_concatenated_grouped['VOIP.INTRA.FREQ.HANDOVER.OUT.SUCCESS.RATE.custom'] = round(100*(df_concatenated_grouped['L.HHO.IntraeNB.IntraFreq.ExecSuccOut.VoIP (None)']+df_concatenated_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut.VoIP (None)'])/(df_concatenated_grouped['L.HHO.IntraeNB.IntraFreq.ExecAttOut.VoIP (None)']+df_concatenated_grouped['L.HHO.IntereNB.IntraFreq.ExecAttOut.VoIP (None)']),2)

    df_concatenated_grouped['INTERFREQ.HANDOVER.OUT.SUCCESS.RATE.custom'] = round(100*(df_concatenated_grouped['L.HHO.IntraeNB.InterFreq.ExecSuccOut (None)']+df_concatenated_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut (None)'])/(df_concatenated_grouped['L.HHO.IntraeNB.InterFreq.ExecAttOut (None)']+df_concatenated_grouped['L.HHO.IntereNB.InterFreq.ExecAttOut (None)']),2)

    df_concatenated_grouped['VOIP.INTERFREQ.HANDOVER.OUT.SUCCESS.RATE.custom'] = round(100*(df_concatenated_grouped['L.HHO.IntraeNB.InterFreq.ExecSuccOut.VoIP (None)']+df_concatenated_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut.VoIP (None)'])/(df_concatenated_grouped['L.HHO.IntraeNB.InterFreq.ExecAttOut.VoIP (None)']+df_concatenated_grouped['L.HHO.IntereNB.InterFreq.ExecAttOut.VoIP (None)']),2)

    df_concatenated_grouped['CSFB.PREPARATION.SUCCESS.RATE.custom'] = round(100*df_concatenated_grouped['L.CSFB.PrepSucc (None)']/df_concatenated_grouped['L.CSFB.PrepAtt (None)'],2)

    df_concatenated_grouped['CSFB.Execution.SUCCESS.RATE.custom'] = round(100*df_concatenated_grouped['L.CSFB.E2W (None)']/df_concatenated_grouped['L.CSFB.PrepSucc (None)'],2)

    df_concatenated_grouped['SRVCC.SUCCESS.RATE.custom'] = round(100*(df_concatenated_grouped['L.IRATHO.SRVCC.E2W.ExecSuccOut (None)']-df_concatenated_grouped['L.IRATHO.SRVCC.E2W.MMEAbnormRsp (None)'])/df_concatenated_grouped['L.IRATHO.SRVCC.E2W.ExecAttOut (None)'],2)

    df_concatenated_grouped['User.DL.Throughput.Protocol'] = round((df_concatenated_grouped['L.Thrp.bits.DL (bit)']-df_concatenated_grouped['L.Thrp.bits.DL.LastTTI (bit)'])/df_concatenated_grouped['L.Thrp.Time.DL.RmvLastTTI (ms)'],2)

    df_concatenated_grouped['User.UL.Throughput.Protocol'] = round((df_concatenated_grouped['L.Thrp.bits.UL (bit)']-df_concatenated_grouped['L.Thrp.bits.UE.UL.LastTTI (bit)'])/df_concatenated_grouped['L.Thrp.Time.UE.UL.RmvLastTTI (ms)'],2)

    df_concatenated_grouped['QCI9.DL.THROUGHPUT.custom'] = round((df_concatenated_grouped['L.Thrp.bits.DL.QCI.9 (bit)']-df_concatenated_grouped['L.Thrp.bits.DL.LastTTI.QCI.9 (bit)'])/(df_concatenated_grouped['L.Thrp.Time.DL.RmvLastTTI.QCI.9 (ms)']*1000),2)

    df_concatenated_grouped['QCI9.UL.THROUGHPUT.custom'] = round(df_concatenated_grouped['L.Thrp.bits.UL.QCI.9 (bit)']/(df_concatenated_grouped['L.Thrp.Time.UL.QCI.9 (ms)']*1000),2)

    df_concatenated_grouped['QCI8.DL.THROUGHPUT.custom'] = round((df_concatenated_grouped['L.Thrp.bits.DL.QCI.8 (bit)']-df_concatenated_grouped['L.Thrp.bits.DL.LastTTI.QCI.8 (bit)'])/(df_concatenated_grouped['L.Thrp.Time.DL.RmvLastTTI.QCI.8 (ms)']*1000),2)

    df_concatenated_grouped['QCI8.UL.THROUGHPUT.custom'] = round(df_concatenated_grouped['L.Thrp.bits.UL.QCI.8 (bit)']/(df_concatenated_grouped['L.Thrp.Time.UL.QCI.8 (ms)']*1000),2)

    df_concatenated_grouped['Cell.DL.throughput'] = round(df_concatenated_grouped['L.Thrp.bits.DL (bit)']/df_concatenated_grouped['L.Thrp.Time.Cell.DL.HighPrecision (ms)'],2)

    df_concatenated_grouped['Cell.UL.throughput'] = round(df_concatenated_grouped['L.Thrp.bits.UL (bit)']/df_concatenated_grouped['L.Thrp.Time.Cell.UL.HighPrecision (ms)'],2)

    df_concatenated_grouped['DL.PACKET.LOSS.RATE.custom'] = round(100*df_concatenated_grouped['L.Traffic.DL.PktUuLoss.Loss (packet)']/df_concatenated_grouped['L.Traffic.DL.PktUuLoss.Tot (packet)'],2)

    df_concatenated_grouped['QCI.1.Service.Downlink.AirInterface.Packet.Loss.Rate.custom'] = round(100*df_concatenated_grouped['L.Traffic.DL.PktUuLoss.Loss.QCI.1 (packet)']/df_concatenated_grouped['L.Traffic.DL.PktUuLoss.Tot.QCI.1 (packet)'],2)

    df_concatenated_grouped['UL.PACKET.LOSS.RATE.custom'] = round(100*df_concatenated_grouped['L.Traffic.UL.PktLoss.Loss (packet)']/df_concatenated_grouped['L.Traffic.UL.PktLoss.Tot (packet)'],2)

    df_concatenated_grouped['QCI.1.Service.Uplink.AirInterface.Packet.Loss.Rate.custom'] = round(100*df_concatenated_grouped['L.Traffic.UL.PktLoss.Loss.QCI.1 (packet)']/df_concatenated_grouped['L.Traffic.UL.PktLoss.Tot.QCI.1 (packet)'],2)

    df_concatenated_grouped['DL.PRB.Usage.Rate'] = round(100*(df_concatenated_grouped['L.ChMeas.PRB.DL.Used.Avg (None)']/df_concatenated_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)

    df_concatenated_grouped['UL.PRB.Usage.Rate'] = round(100*(df_concatenated_grouped['L.ChMeas.PRB.UL.Used.Avg (None)']/df_concatenated_grouped['L.ChMeas.PRB.UL.Avail (None)']),2)
    
    df_concatenated_grouped['VoIP.Traffic.Erlang.custom'] = round((df_concatenated_grouped['L.E-RAB.SessionTime.HighPrecision.QCI1 (100 ms)'])/(36000),2)

    # Final data to return
    return df_concatenated_grouped[['DATE', 'SIGNALING.L.RRC.SETUP.SUCCESS.RATE.custom', 'SERVICE.L.RRC.SETUP.SUCCESS.RATE.custom', 'ERAB.ESTABLISH.SUCCESS.RATE.custom', 'VOIP.ERAB.ESTABLISH.SUCCESS.RATE.custom', 'S1.SIG.CONN.SETUP.SUCCESS.RATE.custom', 'CALL.SETUP.SUCCESS.RATE.custom', 'VOIP.CALL.SETUP.SUCCESS.RATE.custom', 'SERVICE.RETAINABILITY.custom', 'VOIP.CALL.RETAINABILITY.custom', 'INTRA.FREQ.HANDOVER.OUT.SUCCESS.RATE.custom', 'VOIP.INTRA.FREQ.HANDOVER.OUT.SUCCESS.RATE.custom', 'INTERFREQ.HANDOVER.OUT.SUCCESS.RATE.custom', 'VOIP.INTERFREQ.HANDOVER.OUT.SUCCESS.RATE.custom', 'CSFB.PREPARATION.SUCCESS.RATE.custom', 'CSFB.Execution.SUCCESS.RATE.custom', 'SRVCC.SUCCESS.RATE.custom', 'User.DL.Throughput.Protocol', 'User.UL.Throughput.Protocol', 'QCI9.DL.THROUGHPUT.custom', 'QCI9.UL.THROUGHPUT.custom', 'QCI8.DL.THROUGHPUT.custom', 'QCI8.UL.THROUGHPUT.custom', 'Cell.DL.throughput', 'Cell.UL.throughput', 'DL.PACKET.LOSS.RATE.custom', 'QCI.1.Service.Downlink.AirInterface.Packet.Loss.Rate.custom', 'UL.PACKET.LOSS.RATE.custom', 'QCI.1.Service.Uplink.AirInterface.Packet.Loss.Rate.custom', 'L.UL.Interference.Avg (dBm)', 'L.UL.Interference.Max (dBm)', 'L.UL.Interference.Min (dBm)', 'DL.PRB.Usage.Rate', 'UL.PRB.Usage.Rate', 'L.Thrp.bits.DL (bit)', 'L.Thrp.bits.UL (bit)', 'L.Thrp.bits.DL.QCI.9 (bit)', 'L.Thrp.bits.UL.QCI.9 (bit)', 'L.Thrp.bits.DL.QCI.8 (bit)', 'L.Thrp.bits.UL.QCI.8 (bit)', 'VoIP.Traffic.Erlang.custom', 'L.Traffic.User.Avg (None)', 'L.Traffic.ActiveUser.Avg (None)', 'L.ChMeas.CQI.DL.0 (None)', 'L.ChMeas.CQI.DL.1 (None)', 'L.ChMeas.CQI.DL.2 (None)', 'L.ChMeas.CQI.DL.3 (None)', 'L.ChMeas.CQI.DL.4 (None)', 'L.ChMeas.CQI.DL.5 (None)', 'L.ChMeas.CQI.DL.6 (None)', 'L.ChMeas.CQI.DL.7 (None)', 'L.ChMeas.CQI.DL.8 (None)', 'L.ChMeas.CQI.DL.9 (None)', 'L.ChMeas.CQI.DL.10 (None)', 'L.ChMeas.CQI.DL.11 (None)', 'L.ChMeas.CQI.DL.12 (None)', 'L.ChMeas.CQI.DL.13 (None)', 'L.ChMeas.CQI.DL.14 (None)', 'L.ChMeas.CQI.DL.15 (None)']]

def get_cell_bh_kpi(df_concatenated_grouped):
    df_concatenated_grouped['Cell.DL.throughput'] = round(df_concatenated_grouped['L.Thrp.bits.DL (bit)']/df_concatenated_grouped['L.Thrp.Time.Cell.DL.HighPrecision (ms)'],2)

    df_concatenated_grouped['Cell.UL.throughput'] = round(df_concatenated_grouped['L.Thrp.bits.UL (bit)']/df_concatenated_grouped['L.Thrp.Time.Cell.UL.HighPrecision (ms)'],2)
    
    df_concatenated_grouped['DL.PRB.Usage.Rate'] = round(100*(df_concatenated_grouped['L.ChMeas.PRB.DL.Used.Avg (None)']/df_concatenated_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)

    df_concatenated_grouped['UL.PRB.Usage.Rate'] = round(100*(df_concatenated_grouped['L.ChMeas.PRB.UL.Used.Avg (None)']/df_concatenated_grouped['L.ChMeas.PRB.UL.Avail (None)']),2)
    
    # Final data to return
    return df_concatenated_grouped[['DATE', 'Cell.DL.throughput', 'Cell.UL.throughput','DL.PRB.Usage.Rate', 'UL.PRB.Usage.Rate', 'L.Traffic.User.Avg (None)', 'L.Traffic.ActiveUser.Avg (None)']]

def get_mocn_day_kpi(df_grouped):
    df_grouped['MOCNTraffic.custom'] = round((df_grouped['L.Thrp.bits.DL.PLMN (bit)']+df_grouped['L.Thrp.bits.UL.PLMN (bit)'])/(8*1024*1024*1024),2)

    df_grouped['VoLTETraffic.custom'] = round(df_grouped['L.E-RAB.SessionTime.HighPrecision.PLMN.QCI1 (100 ms)']/36000,2)

    df_grouped['DLTraffic.custom'] = round(df_grouped['L.Thrp.bits.DL.PLMN (bit)']/(8*1024*1024*1024),2)

    df_grouped['ULTraffic.custom'] = round(df_grouped['L.Thrp.bits.UL.PLMN (bit)']/(8*1024*1024*1024),2)

    df_grouped['WBBTraffic.custom'] = round((df_grouped['L.Thrp.bits.DL.PLMN.QCI.8 (bit)']+df_grouped['L.Thrp.bits.UL.PLMN.QCI.8 (bit)'])/(8*1024*1024*1024),2)

    df_grouped['Accessibility.custom'] = round(100*df_grouped['L.E-RAB.SuccEst.PLMN (None)']/df_grouped['L.E-RAB.AttEst.PLMN (None)'],2)

    df_grouped['VoLTEAcc.custom'] = round(100*df_grouped['L.E-RAB.SuccEst.PLMN.QCI.1 (None)']/df_grouped['L.E-RAB.AttEst.PLMN.QCI.1 (None)'],2)

    df_grouped['Retainability.custom'] = round(100*df_grouped['L.E-RAB.AbnormRel.PLMN (None)']/(df_grouped['L.E-RAB.AbnormRel.PLMN (None)']+df_grouped['L.E-RAB.NormRel.PLMN (None)']+df_grouped['L.IRATHO.E2W.ExecSuccOut.PLMN (None)']),2)

    df_grouped['RetainabilityMME.custom'] = round(100*(df_grouped['L.E-RAB.AbnormRel.PLMN (None)']+df_grouped['L.E-RAB.AbnormRel.MME.PLMN (None)'])/(df_grouped['L.E-RAB.AbnormRel.PLMN (None)']+df_grouped['L.E-RAB.NormRel.PLMN (None)']+df_grouped['L.IRATHO.E2W.ExecSuccOut.PLMN (None)']),2)

    df_grouped['VoLTERet.custom'] = round(100*df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)']/(df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)']+df_grouped['L.E-RAB.NormRel.PLMN.QCI.1 (None)']),2)

    df_grouped['VoLTERetMME.custom'] = round(100*(df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)']+df_grouped['L.E-RAB.AbnormRel.MME.VoIP.PLMN (None)'])/(df_grouped['L.E-RAB.AbnormRel.PLMN.QCI.1 (None)']+df_grouped['L.E-RAB.NormRel.PLMN.QCI.1 (None)']),2)

    df_grouped['DLUserThroughput.custom'] = round((df_grouped['L.Thrp.bits.DL.PLMN (bit)']-df_grouped['L.Thrp.bits.DL.LastTTI.PLMN (bit)'])/(df_grouped['L.Thrp.Time.DL.RmvLastTTI.PLMN (ms)']*1000),2)

    df_grouped['ULUserThroughput.custom'] = round((df_grouped['L.Thrp.bits.UL.PLMN (bit)']-df_grouped['L.Thrp.bits.UE.UL.LastTTI.PLMN (bit)'])/(df_grouped['L.Thrp.Time.UE.UL.RmvLastTTI.PLMN (ms)']*1000),2)

    df_grouped['DLPRB.custom'] = round(100*(df_grouped['L.RBUsedOwn.DL.PLMN (None)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)

    df_grouped['ULPRB.custom'] = round(100*(df_grouped['L.RBUsedOwn.UL.PLMN (None)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)

    df_grouped['HOIntraSuccessRate.custom'] = round(100*(df_grouped['L.HHO.IntraeNB.IntraFreq.ExecSuccOut.PLMN (None)']+df_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.IntraeNB.IntraFreq.PrepAttOut.PLMN (None)']+df_grouped['L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)']),2)

    df_grouped['HOInterSuccessRate.custom'] = round(100*(df_grouped['L.HHO.IntraeNB.InterFreq.ExecSuccOut.PLMN (None)']+df_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.IntraeNB.InterFreq.PrepAttOut.PLMN (None)']+df_grouped['L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)']),2)

    df_grouped['HOX2SuccessRate.custom'] = round(100*(df_grouped['L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)']+df_grouped['L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)']+df_grouped['L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)']),2)

    df_grouped['HOS1SuccessRate.custom'] = round(100*(df_grouped['L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)']-df_grouped['L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)']+df_grouped['L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)']-df_grouped['L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)'])/(df_grouped['L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)']-df_grouped['L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)']+df_grouped['L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)']-df_grouped['L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)']),2)

    df_grouped['CSFBSuccessRate.custom'] = round(100*df_grouped['L.CSFB.PrepSucc.PLMN (None)']/df_grouped['L.CSFB.PrepAtt.PLMN (None)'],2)

    # Final data to return
    return df_grouped[['DATE', 'CnOperator', 'L.Traffic.User.Avg.PLMN (None)', 'MOCNTraffic.custom','VoLTETraffic.custom', 'DLTraffic.custom', 'ULTraffic.custom', 'WBBTraffic.custom', 'Accessibility.custom', 'VoLTEAcc.custom', 'Retainability.custom', 'RetainabilityMME.custom', 'VoLTERet.custom', 'VoLTERetMME.custom', 'DLUserThroughput.custom', 'ULUserThroughput.custom', 'DLPRB.custom', 'ULPRB.custom', 'HOIntraSuccessRate.custom', 'HOInterSuccessRate.custom', 'HOX2SuccessRate.custom', 'HOS1SuccessRate.custom', 'CSFBSuccessRate.custom']]
    
def get_mocn_bh_kpi(df_grouped):
    df_grouped['DLPRB.custom'] = round(100*(df_grouped['L.RBUsedOwn.DL.PLMN (None)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)

    df_grouped['ULPRB.custom'] = round(100*(df_grouped['L.RBUsedOwn.UL.PLMN (None)'])/(df_grouped['L.ChMeas.PRB.DL.Avail (None)']),2)

    # Final data to return
    return df_grouped[['DATE', 'CnOperator', 'L.Traffic.User.Avg.PLMN (None)', 'DLPRB.custom', 'ULPRB.custom']] 

class LTEReportCalc():

    def __init__(self, df, counters_type) -> None:
        if counters_type == 'CELL':
            self.lteCellCounters_df = df.copy()
            self.lteCellCounters_df['Start Time'] = self.lteCellCounters_df['Start Time'].str.replace(':00 DST', '')
            self.lteCellCounters_df['Start Time'] = self.lteCellCounters_df['Start Time'].str.replace(':00:00', ':00')
            self.lteCellCounters_df['Start Time'] = pd.to_datetime(self.lteCellCounters_df['Start Time'])
            self.lteCellCounters_df['DATE'] = pd.to_datetime(self.lteCellCounters_df['Start Time']).dt.date
            self.lteCellCounters_df['CellName'] = self.lteCellCounters_df['Cell'].str.split(pat="=", expand=True)[3].str.split(pat=",", expand=True)[0]
            self.lteCellCounters_df = self.lteCellCounters_df.drop(['Cell'], axis=1)
        elif counters_type == 'MOCN':
            self.lteMocnCounters_df = df.copy()
            self.lteMocnCounters_df['Start Time'] = self.lteMocnCounters_df['Start Time'].str.replace(':00 DST', '')
            self.lteMocnCounters_df['Start Time'] = self.lteMocnCounters_df['Start Time'].str.replace(':00:00', ':00')
            self.lteMocnCounters_df['Start Time'] = pd.to_datetime(self.lteMocnCounters_df['Start Time'])
            self.lteMocnCounters_df['DATE'] = pd.to_datetime(self.lteMocnCounters_df['Start Time']).dt.date
            self.lteMocnCounters_df['CellName'] = self.lteMocnCounters_df['Cell'].str.split(pat="=", expand=True)[3].str.split(pat=",", expand=True)[0]
            self.lteMocnCounters_df = self.lteMocnCounters_df.drop(['Cell'], axis=1)
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=0, Mobile Country Code=334, Mobile Network Code=090', 'ATT')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=1, Mobile Country Code=334, Mobile Network Code=050', 'ATT')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=2, Mobile Country Code=334, Mobile Network Code=03', 'TLF')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=0, Object ID=0', 'ATT')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=1, Object ID=0', 'ATT')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=2, Object ID=0', 'TLF')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=0', 'ATT')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=1', 'ATT')
            self.lteMocnCounters_df['CnOperator'] = self.lteMocnCounters_df['CnOperator'].str.replace('CN Operator ID=2', 'TLF')

    def add_prb_avail_mocn(self, prb_avail_df):
        self.lteMocnCounters_df = self.lteMocnCounters_df.merge(prb_avail_df, on='CellName', how='left')

    def group_data1(self, by='ALL_DAY'):
        if by == 'ALL_DAY':
            group_cell_by = ['DATE']
        elif by == 'BH':
            group_cell_by = ['DATE', 'Start Time', 'CellName']

        self.lteCellCounters_df = self.lteCellCounters_df.groupby(by=group_cell_by).agg({
                'L.Thrp.bits.DL (bit)': 'sum',
                'L.Thrp.Time.Cell.DL.HighPrecision (ms)': 'sum',
                'L.Thrp.bits.UL (bit)': 'sum',
                'L.Thrp.Time.Cell.UL.HighPrecision (ms)': 'sum',
                'L.ChMeas.PRB.DL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.DL.Avail (None)': 'sum',
                'L.ChMeas.PRB.UL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.UL.Avail (None)': 'sum',
                'L.Traffic.ActiveUser.Avg (None)': 'sum',
                'L.Traffic.User.Avg (None)': 'sum',
                'L.RRC.ConnReq.Succ.MoSig (None)': 'sum',
                'L.RRC.ConnReq.Att.MoSig (None)': 'sum',
                'L.RRC.ConnReq.Succ.DelayTol (None)': 'sum',
                'L.RRC.ConnReq.Succ.Emc (None)': 'sum',
                'L.RRC.ConnReq.Succ.HighPri (None)': 'sum',
                'L.RRC.ConnReq.Succ.MoData (None)': 'sum',
                'L.RRC.ConnReq.Succ.MoVoiceCall (None)': 'sum',
                'L.RRC.ConnReq.Succ.Mt (None)': 'sum',
                'L.RRC.ConnReq.Att.DelayTol (None)': 'sum',
                'L.RRC.ConnReq.Att.Emc (None)': 'sum',
                'L.RRC.ConnReq.Att.HighPri (None)': 'sum',
                'L.RRC.ConnReq.Att.MoData (None)': 'sum',
                'L.RRC.ConnReq.Att.MoVoiceCall (None)': 'sum',
                'L.RRC.ConnReq.Att.Mt (None)': 'sum',
                'L.E-RAB.SuccEst.QCI.1 (None)': 'sum',
                'L.E-RAB.AttEst.QCI.1 (None)': 'sum',
                'L.E-RAB.FailEst.X2AP.VoIP (None)': 'sum',
                'L.E-RAB.SuccEst (None)': 'sum',
                'L.E-RAB.AttEst (None)': 'sum',
                'L.E-RAB.FailEst.X2AP (None)': 'sum',
                'L.S1Sig.ConnEst.Succ (None)': 'sum',
                'L.S1Sig.ConnEst.Att (None)': 'sum',
                'L.E-RAB.AbnormRel (None)': 'sum',
                'L.E-RAB.NormRel (None)': 'sum',
                'L.E-RAB.NormRel.IRatHOOut (None)': 'sum',
                'L.E-RAB.AbnormRel.QCI.1 (None)': 'sum',
                'L.E-RAB.NormRel.QCI.1 (None)': 'sum',
                'L.E-RAB.NormRel.IRatHOOut.QCI.1 (None)': 'sum',
                'L.HHO.IntraeNB.IntraFreq.ExecSuccOut (None)': 'sum',
                'L.HHO.IntereNB.IntraFreq.ExecSuccOut (None)': 'sum',
                'L.HHO.IntraeNB.IntraFreq.ExecAttOut (None)': 'sum',
                'L.HHO.IntereNB.IntraFreq.ExecAttOut (None)': 'sum',
                'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.VoIP (None)': 'sum',
                'L.HHO.IntereNB.IntraFreq.ExecSuccOut.VoIP (None)': 'sum',
                'L.HHO.IntraeNB.IntraFreq.ExecAttOut.VoIP (None)': 'sum',
                'L.HHO.IntereNB.IntraFreq.ExecAttOut.VoIP (None)': 'sum',
                'L.HHO.IntraeNB.InterFreq.ExecSuccOut (None)': 'sum',
                'L.HHO.IntereNB.InterFreq.ExecSuccOut (None)': 'sum',
                'L.HHO.IntraeNB.InterFreq.ExecAttOut (None)': 'sum',
                'L.HHO.IntereNB.InterFreq.ExecAttOut (None)': 'sum',
                'L.HHO.IntraeNB.InterFreq.ExecSuccOut.VoIP (None)': 'sum',
                'L.HHO.IntereNB.InterFreq.ExecSuccOut.VoIP (None)': 'sum',
                'L.HHO.IntraeNB.InterFreq.ExecAttOut.VoIP (None)': 'sum',
                'L.HHO.IntereNB.InterFreq.ExecAttOut.VoIP (None)': 'sum',
                'L.CSFB.PrepSucc (None)': 'sum',
                'L.CSFB.PrepAtt (None)': 'sum',
                'L.CSFB.E2W (None)': 'sum',
                'L.IRATHO.SRVCC.E2W.ExecSuccOut (None)': 'sum',
                'L.IRATHO.SRVCC.E2W.MMEAbnormRsp (None)': 'sum',
                'L.IRATHO.SRVCC.E2W.ExecAttOut (None)': 'sum',
                'L.Thrp.bits.DL.LastTTI (bit)': 'sum',
                'L.Thrp.Time.DL.RmvLastTTI (ms)': 'sum',
                'L.Thrp.bits.UE.UL.LastTTI (bit)': 'sum',
                'L.Thrp.Time.UE.UL.RmvLastTTI (ms)': 'sum',
                'L.Thrp.bits.DL.QCI.9 (bit)': 'sum',
                'L.Thrp.bits.DL.LastTTI.QCI.9 (bit)': 'sum',
                'L.Thrp.Time.DL.RmvLastTTI.QCI.9 (ms)': 'sum',
                'L.Thrp.bits.UL.QCI.9 (bit)': 'sum',
                'L.Thrp.Time.UL.QCI.9 (ms)': 'sum',
                'L.Thrp.bits.DL.QCI.8 (bit)': 'sum',
                'L.Thrp.bits.DL.LastTTI.QCI.8 (bit)': 'sum',
                'L.Thrp.Time.DL.RmvLastTTI.QCI.8 (ms)': 'sum',
                'L.Thrp.bits.UL.QCI.8 (bit)': 'sum',
                'L.Thrp.Time.UL.QCI.8 (ms)': 'sum',
                'L.Traffic.DL.PktUuLoss.Loss (packet)': 'sum',
                'L.Traffic.DL.PktUuLoss.Tot (packet)': 'sum',
                'L.Traffic.DL.PktUuLoss.Loss.QCI.1 (packet)': 'sum',
                'L.Traffic.DL.PktUuLoss.Tot.QCI.1 (packet)': 'sum',
                'L.Traffic.UL.PktLoss.Loss (packet)': 'sum',
                'L.Traffic.UL.PktLoss.Tot (packet)': 'sum',
                'L.Traffic.UL.PktLoss.Loss.QCI.1 (packet)': 'sum',
                'L.Traffic.UL.PktLoss.Tot.QCI.1 (packet)': 'sum',
                'L.E-RAB.SessionTime.HighPrecision.QCI1 (100 ms)': 'sum'
            })
        if by == 'BH':
            # Sort PRB usage descending and then remove duplicates for each date/cell/cnoperator
            # Cell data
            self.lteCellCounters_df = self.lteCellCounters_df.reset_index()
            self.lteCellCounters_df['DL_PRB_USAGE_FOR_BH'] = round(100*(self.lteCellCounters_df['L.ChMeas.PRB.DL.Used.Avg (None)']/self.lteCellCounters_df['L.ChMeas.PRB.DL.Avail (None)']),2)
            self.lteCellCounters_df = self.lteCellCounters_df.sort_values('DL_PRB_USAGE_FOR_BH', ascending=False).drop_duplicates(['DATE', 'CellName'])
            self.lteCellCounters_df = self.lteCellCounters_df.drop(['DL_PRB_USAGE_FOR_BH'], axis=1)
            # self.lteCellCounters_df = self.lteCellCounters_df.reset_index()
            self.lteCellCounters_df = self.lteCellCounters_df.groupby(by=['DATE']).agg({
                'L.Thrp.bits.DL (bit)': 'sum',
                'L.Thrp.Time.Cell.DL.HighPrecision (ms)': 'sum',
                'L.Thrp.bits.UL (bit)': 'sum',
                'L.Thrp.Time.Cell.UL.HighPrecision (ms)': 'sum',
                'L.ChMeas.PRB.DL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.DL.Avail (None)': 'sum',
                'L.ChMeas.PRB.UL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.UL.Avail (None)': 'sum',
                'L.Traffic.ActiveUser.Avg (None)': 'sum',
                'L.Traffic.User.Avg (None)': 'sum'
            })
        self.lteCellCounters_df = self.lteCellCounters_df.reset_index()

    def group_data2(self, by='ALL_DAY'):
        if by == 'ALL_DAY':
            group_cell_by = ['DATE']
        elif by == 'BH':
            group_cell_by = ['DATE', 'Start Time', 'CellName']

        self.lteCellCounters_df = self.lteCellCounters_df.groupby(by=group_cell_by).agg({
            'L.UL.Interference.Avg (dBm)': 'mean',
            'L.UL.Interference.Max (dBm)': 'mean',
            'L.UL.Interference.Min (dBm)': 'mean',
            'L.ChMeas.CQI.DL.0 (None)': 'sum',
            'L.ChMeas.CQI.DL.1 (None)': 'sum',
            'L.ChMeas.CQI.DL.10 (None)': 'sum',
            'L.ChMeas.CQI.DL.11 (None)': 'sum',
            'L.ChMeas.CQI.DL.12 (None)': 'sum',
            'L.ChMeas.CQI.DL.13 (None)': 'sum',
            'L.ChMeas.CQI.DL.14 (None)': 'sum',
            'L.ChMeas.CQI.DL.15 (None)': 'sum',
            'L.ChMeas.CQI.DL.2 (None)': 'sum',
            'L.ChMeas.CQI.DL.3 (None)': 'sum',
            'L.ChMeas.CQI.DL.4 (None)': 'sum',
            'L.ChMeas.CQI.DL.5 (None)': 'sum',
            'L.ChMeas.CQI.DL.6 (None)': 'sum',
            'L.ChMeas.CQI.DL.7 (None)': 'sum',
            'L.ChMeas.CQI.DL.8 (None)': 'sum',
            'L.ChMeas.CQI.DL.9 (None)': 'sum'
        })
        if by == 'BH':
            # Sort PRB usage descending and then remove duplicates for each date/cell/cnoperator
            # Cell data
            self.lteCellCounters_df = self.lteCellCounters_df.reset_index()
            self.lteCellCounters_df['DL_PRB_USAGE_FOR_BH'] = round(100*(self.lteCellCounters_df['L.ChMeas.PRB.DL.Used.Avg (None)']/self.lteCellCounters_df['L.ChMeas.PRB.DL.Avail (None)']),2)
            self.lteCellCounters_df = self.lteCellCounters_df.sort_values('DL_PRB_USAGE_FOR_BH', ascending=False).drop_duplicates(['DATE', 'CellName'])
            self.lteCellCounters_df = self.lteCellCounters_df.drop(['DL_PRB_USAGE_FOR_BH'], axis=1)
            self.lteCellCounters_df = self.lteCellCounters_df.reset_index()
            self.lteCellCounters_df = self.lteCellCounters_df.groupby(by=['DATE']).agg({
                'L.Thrp.bits.DL (bit)': 'sum',
                'L.Thrp.Time.Cell.DL.HighPrecision (ms)': 'sum',
                'L.Thrp.bits.UL (bit)': 'sum',
                'L.Thrp.Time.Cell.UL.HighPrecision (ms)': 'sum',
                'L.ChMeas.PRB.DL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.DL.Avail (None)': 'sum',
                'L.ChMeas.PRB.UL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.UL.Avail (None)': 'sum',
                'L.Traffic.ActiveUser.Avg (None)': 'sum',
                'L.Traffic.User.Avg (None)': 'sum'
            })
        self.lteCellCounters_df = self.lteCellCounters_df.reset_index()

    def group_data_mocn(self, by='ALL_DAY'):
        if by == 'ALL_DAY':
            group_mocn_by = ['DATE', 'CnOperator']
        elif by == 'BH':
            group_mocn_by = ['DATE', 'Start Time', 'CellName', 'CnOperator']

        self.lteMocnCounters_df = self.lteMocnCounters_df.groupby(by=group_mocn_by).agg({
            'L.Traffic.User.Avg.PLMN (None)': 'sum',
            'L.Thrp.bits.DL.PLMN (bit)': 'sum',
            'L.Thrp.bits.UL.PLMN (bit)': 'sum',
            'L.E-RAB.SessionTime.HighPrecision.PLMN.QCI1 (100 ms)': 'sum',
            'L.Thrp.bits.DL.PLMN.QCI.8 (bit)': 'sum',
            'L.Thrp.bits.UL.PLMN.QCI.8 (bit)': 'sum',
            'L.E-RAB.SuccEst.PLMN (None)': 'sum',
            'L.E-RAB.AttEst.PLMN (None)': 'sum',
            'L.E-RAB.SuccEst.PLMN.QCI.1 (None)': 'sum',
            'L.E-RAB.AttEst.PLMN.QCI.1 (None)': 'sum',
            'L.E-RAB.AbnormRel.PLMN (None)': 'sum',
            'L.E-RAB.NormRel.PLMN (None)': 'sum',
            'L.IRATHO.E2W.ExecSuccOut.PLMN (None)': 'sum',
            'L.E-RAB.AbnormRel.MME.PLMN (None)': 'sum',
            'L.E-RAB.AbnormRel.PLMN.QCI.1 (None)': 'sum',
            'L.E-RAB.NormRel.PLMN.QCI.1 (None)': 'sum',
            'L.E-RAB.AbnormRel.MME.VoIP.PLMN (None)': 'sum',
            'L.Thrp.bits.DL.LastTTI.PLMN (bit)': 'sum',
            'L.Thrp.Time.DL.RmvLastTTI.PLMN (ms)': 'sum',
            'L.Thrp.bits.UE.UL.LastTTI.PLMN (bit)': 'sum',
            'L.Thrp.Time.UE.UL.RmvLastTTI.PLMN (ms)': 'sum',
            'L.RBUsedOwn.DL.PLMN (None)': 'mean',
            'L.RBUsedOwn.UL.PLMN (None)': 'mean',
            'L.HHO.IntraeNB.IntraFreq.ExecSuccOut.PLMN (None)': 'sum',
            'L.HHO.IntereNB.IntraFreq.ExecSuccOut.PLMN (None)': 'sum',
            'L.HHO.IntraeNB.IntraFreq.PrepAttOut.PLMN (None)': 'sum',
            'L.HHO.IntereNB.IntraFreq.PrepAttOut.PLMN (None)': 'sum',
            'L.HHO.IntraeNB.InterFreq.ExecSuccOut.PLMN (None)': 'sum',
            'L.HHO.IntereNB.InterFreq.ExecSuccOut.PLMN (None)': 'sum',
            'L.HHO.IntraeNB.InterFreq.PrepAttOut.PLMN (None)': 'sum',
            'L.HHO.IntereNB.InterFreq.PrepAttOut.PLMN (None)': 'sum',
            'L.HHO.X2.IntraFreq.ExecSuccOut.PLMN (None)': 'sum',
            'L.HHO.X2.InterFreq.ExecSuccOut.PLMN (None)': 'sum',
            'L.HHO.X2.IntraFreq.PrepAttOut.PLMN (None)': 'sum',
            'L.HHO.X2.InterFreq.PrepAttOut.PLMN (None)': 'sum',
            'L.CSFB.PrepSucc.PLMN (None)': 'sum',
            'L.CSFB.PrepAtt.PLMN (None)': 'sum',
            'L.ChMeas.PRB.DL.Avail (None)': 'mean'
        })
        if by == 'BH':
            # Sort PRB usage descending and then remove duplicates for each date/cell/cnoperator
            # MOCN data
            self.lteMocnCounters_df = self.lteMocnCounters_df.reset_index()            
            self.lteMocnCounters_df['DL_PRB_USAGE_FOR_BH'] = round(100*(self.lteMocnCounters_df['L.RBUsedOwn.DL.PLMN (None)']/self.lteMocnCounters_df['L.ChMeas.PRB.DL.Avail (None)']),2)
            self.lteMocnCounters_df = self.lteMocnCounters_df.sort_values('DL_PRB_USAGE_FOR_BH', ascending=False).drop_duplicates(['DATE', 'CellName','CnOperator'])
            self.lteMocnCounters_df = self.lteMocnCounters_df.drop(['DL_PRB_USAGE_FOR_BH'], axis=1)
            self.lteMocnCounters_df = self.lteMocnCounters_df.reset_index()
            self.lteMocnCounters_df =self.lteMocnCounters_df.groupby(by=['DATE', 'CnOperator']).agg({
                'L.Traffic.User.Avg.PLMN (None)': 'sum',
                'L.RBUsedOwn.DL.PLMN (None)': 'mean',
                'L.RBUsedOwn.UL.PLMN (None)': 'mean',
                'L.ChMeas.PRB.DL.Avail (None)': 'mean'
            })
        self.lteMocnCounters_df = self.lteMocnCounters_df.reset_index()

    def group_per_cell(self, by='ALL_DAY'):
        if by == 'BH':
            self.lteCellCounters_df = self.lteCellCounters_df.groupby(by=['DATE', 'Start Time', 'NE Name', 'CellName']).agg({
                'L.Thrp.bits.DL (bit)': 'sum',
                'L.Thrp.bits.DL.LastTTI (bit)': 'sum',
                'L.Thrp.Time.DL.RmvLastTTI (ms)': 'sum',
                'L.ChMeas.PRB.DL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.DL.Avail (None)': 'sum'
            })
        elif by == "ALL_DAY":
            self.lteCellCounters_df = self.lteCellCounters_df.groupby(by=['DATE', 'NE Name', 'CellName']).agg({
                'L.Thrp.bits.DL (bit)': 'sum',
                'L.Thrp.bits.DL.LastTTI (bit)': 'sum',
                'L.Thrp.Time.DL.RmvLastTTI (ms)': 'sum',
                'L.ChMeas.PRB.DL.Used.Avg (None)': 'sum',
                'L.ChMeas.PRB.DL.Avail (None)': 'sum'
            })
        self.lteCellCounters_df = self.lteCellCounters_df.reset_index()
        self.lteCellCounters_df['BH.DL.PRB.Usage'] = round(100*(self.lteCellCounters_df['L.ChMeas.PRB.DL.Used.Avg (None)']/self.lteCellCounters_df['L.ChMeas.PRB.DL.Avail (None)']),2)
        if by == 'BH':
            self.lteCellCounters_df = self.lteCellCounters_df.sort_values('BH.DL.PRB.Usage', ascending=False).drop_duplicates(['DATE', 'CellName'])
        self.lteCellCounters_df['User.DL.Throughput.Protocol'] = round((self.lteCellCounters_df['L.Thrp.bits.DL (bit)']-self.lteCellCounters_df['L.Thrp.bits.DL.LastTTI (bit)'])/self.lteCellCounters_df['L.Thrp.Time.DL.RmvLastTTI (ms)'],2)
        self.lteCellCounters_df = self.lteCellCounters_df[['DATE', 'NE Name', 'CellName', 'BH.DL.PRB.Usage','User.DL.Throughput.Protocol', 'L.Thrp.bits.DL (bit)']]
        
class LTEReportGenerator():
    
    @staticmethod
    def generate_report_data(df_1, df_2, mocn):
        report_writer = ReportWriter()
        df1 = LTEReportCalc(df_1, "CELL")
        NE_check_df1 = count_per_NE(df1.lteCellCounters_df, 'NE Name')
        Cell_check_df1 = count_per_NE(df1.lteCellCounters_df, 'CellName')
        df2 = LTEReportCalc(df_2, "CELL")
        del df_2
        NE_check_df2 = count_per_NE(df2.lteCellCounters_df, 'NE Name')
        Cell_check_df2 = count_per_NE(df2.lteCellCounters_df, 'CellName')
        # Get in separate DF the data about PRB avail to include in MOCN data
        prb_avail_df = df1.lteCellCounters_df[['CellName', 'L.ChMeas.PRB.DL.Avail (None)']].drop_duplicates(subset=['CellName'])
        df1.group_data1()
        df2.group_data2()
        # Concatenate counters 1 and 2 DF to 1
        lteCounters_data_group_df = pd.concat([df1.lteCellCounters_df, df2.lteCellCounters_df], axis=1, join="inner")
        print('Analysis started...')
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
        del NE_check_df1, NE_check_df2, Cell_check_df1, Cell_check_df2
        print('NE Check for cluster finished...')
        # All day for cluster 
        report_writer.create_sheet('LTE All Day - Cluster')
        ws = report_writer.get_sheet('LTE All Day - Cluster')
        kpi_df = get_cell_day_kpi(lteCounters_data_group_df)
        report_writer.save_df_sheet(ws, kpi_df)
        print('All day for cluster finished...')
        # BH data for cluster
        df1 = LTEReportCalc(df_1, "CELL")
        df1.group_data1(by="BH")
        report_writer.create_sheet('LTE BH - Cluster')
        ws = report_writer.get_sheet('LTE BH - Cluster')
        kpi_df = get_cell_bh_kpi(df1.lteCellCounters_df)
        report_writer.save_df_sheet(ws, kpi_df)
        print('BH for cluster finished...')
        # All Day per cell
        df1 = LTEReportCalc(df_1, "CELL")
        df1.group_per_cell()
        report_writer.create_sheet('LTE All Day - Cell')
        ws = report_writer.get_sheet('LTE All Day - Cell')
        report_writer.save_df_sheet(ws, df1.lteCellCounters_df)
        print('All Day for cells finished...')
        # BH per cell
        df1 = LTEReportCalc(df_1, "CELL")
        del df_1
        df1.group_per_cell(by="BH")
        report_writer.create_sheet('LTE BH - Cell')
        ws = report_writer.get_sheet('LTE BH - Cell')
        report_writer.save_df_sheet(ws, df1.lteCellCounters_df)
        print('BH for cells finished...')
        # NE Check for MOCN
        df = LTEReportCalc(mocn, "MOCN")
        NE_check_df_mocn = count_per_NE(df.lteMocnCounters_df, 'NE Name')
        report_writer.create_sheet('Data Check MOCN')
        ws = report_writer.get_sheet('Data Check MOCN')
        report_writer.save_df_sheet(ws, NE_check_df_mocn)
        print('NE Check for MOCN finished...')
        # All day for MOCN
        df.add_prb_avail_mocn(prb_avail_df)
        df.group_data_mocn()
        report_writer.create_sheet('LTE All Day - MOCN')
        ws = report_writer.get_sheet('LTE All Day - MOCN')
        kpi_df = get_mocn_day_kpi(df.lteMocnCounters_df)
        report_writer.save_df_sheet(ws, kpi_df)
        print('All day for MOCN finished...')
        # BH for MOCN
        df = LTEReportCalc(mocn, "MOCN")
        df.add_prb_avail_mocn(prb_avail_df)
        df.group_data_mocn(by="BH")
        report_writer.create_sheet('LTE BH - MOCN')
        ws = report_writer.get_sheet('LTE BH - MOCN')
        kpi_df = get_mocn_bh_kpi(df.lteMocnCounters_df)
        report_writer.save_df_sheet(ws, kpi_df)
        print('BH for MOCN finished...')
        report_writer.save_excel_report("LTE")
        print('FINISHED')