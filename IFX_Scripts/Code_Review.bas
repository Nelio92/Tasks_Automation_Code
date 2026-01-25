Attribute VB_Name = "VBT_50_Txge"
'<AUTOCOMMENT:0.4:TOPLEVEL
'.......10........20........30........40........50........60........70..   (max. line width is 120)    ....110.......120
'''@brief TXGE Module performs TXRF PLD and TXLO PLD VLD trimming. Also TX current consumption measurements.
'''@details
'''Test functions for TXRF PLD trimming and TXLO PLD VLD trimming are using the generic CTRX trimming class function
''' to perform the necessary sweeps. The test function for the current consumption measurement requires that the
''' voltage supplies be calibrated to allow for the high current consumption when each PA is enabled. It is expected
''' that the current consumption tests are possible with pairs of TX channels enabled. It is not going to be possible
''' to enable all 8 TX PA channels as the current demand is too high.
'''
'''@author Channon Andrew (PSS SC D RAD PTE TE3 External) <Channon.external@infineon.com>
'''
'''@copyright &copy; Infineon Technologies AG 2025
'>AUTOCOMMENT:0.4:TOPLEVEL

Option Explicit

' CONSTANTS
Private Const TX_TRIM_ERR_0 = 0
Private Const LOIN = 1
Private Const LOOUT = 2
Private Const TXPLD_CH15 = 1
Private Const TXPLD_CH26 = 2
Private Const TXPLD_CH37 = 3
Private Const TXPLD_CH48 = 4
Private Const PLD_BIG_JUMPS_ALLOWED   As Boolean = True
Private Const PLD_TX_START_VAL As Double = 0
Private Const PLD_TX_END_VAL As Double = 7
Private Const PLD_TX_DEFAULT_VAL As Double = 3
Private Const PLD_TX_TARGET_VAL As Double = 0.53 * V
Private Const VLD_TXLO_TARGET_VAL As Double = 0.325 * V
Private Const PLD_TX_AVG_CNT_VAL As Double = HW_AVERAGE_32
Private Const PLD_TX_DECENDING = False
Private Const VLD_TX_END_VAL As Double = 15
Private Const VLD_TX_DEFAULT_VAL As Double = 8
Private Const TRIM_TXLO_TC1_SLOPE = &H4
Private Const TRIM_TXLO_TC2_SLOPE = &H1
Private Const TRIM_TXLO_TC3_SLOPE = &H3
Private Const TRIM_TXLO_TC4_SLOPE = &H3
Private Const TRIM_TXLO_TC5_SLOPE = &H3
Private Const DCVI_DEFAULT_CURRENT As Double = 0 * A
Private Const DCVI_DEFAULT_VOLTAGE As Double = 0 * V
Private Const FORCE_V_NMOS_VAL As Double = 0.3 * V
''' Lo1 test codes for the Buffer sweep. These are the same values as used in the 8191B
Private Const LO1TST_BUF_DAC_CODE = "1,21,31,40,47,54,61,68,75,81,88,94,99,105,112,118,124,131,138,144,150,157,163" & _
                                    ",170,177,185,195,205,215,225,239,255"
Private Const LO1TST_CNT As Double = 31
Private Const VLD_TARGET_VALUE As Double = 25 * mV
Private Const TST_BUF_START_VAL As Double = 1
Private Const TST_BUF_END_VAL As Double = 255

Private Const WAIT_TOTSYS_SETTLE = 2 * ms
Private Const WAIT_ISD_PA_SETTLE = 2 * ms
' MODULE AND DATALOGGING VARIABLES

' Module related
Private mdDcviWait As Double
''' Variables for Measured Txge supply currents and DC currents at the AMUX
Private mwVddCorrMin As New DSPWave
Private mwVddCorrMax As New DSPWave
Private msPatHw As String
Private msPinsCurrCons As String
Private msVdd1v8Tx As String
Private mpldaIddMaxCorner(16) As New PinListData
Private mpldaIddMinCorner(16) As New PinListData
Private mpldaVddMaxCorner(15) As New PinListData
Private mpldaVddMinCorner(15) As New PinListData
Private mpldaCurrResults(58) As New PinListData
Private msdIccVdd1V0MaxCornerSum As New SiteDouble
Private msdIccVdd1V0MinCornerSum As New SiteDouble
Private mpldDpllSense1V2 As New PinListData
Private mpldaTxrfIbgCurr(15) As New PinListData
Private mpldaTxBiasCurr(3) As New PinListData
Private mpldaTxloBiasCurr(4) As New PinListData
Private mpldaDpllBiasCurr(5) As New PinListData
''' Variables for Measured TxgeTxlo DC currents and voltages
Private mpldVldInFqm2Offset As New PinListData
Private mpldVldOutFqm2Offset As New PinListData
Private mpldVldInFqm3Offset As New PinListData
Private mpldVldOutFqm3Offset As New PinListData
Private mpldVldLoin1Offset As New PinListData
Private mpldVldLoin2Offset As New PinListData
Private mpldVldLoout1Offset As New PinListData
Private mpldVldLoout2Offset As New PinListData
Private mpldaTxloVddPost(1) As New PinListData
Private mpldIccStandalone As New PinListData
Private mpldIccTxmonOff As New PinListData
Private mpldIccLoDistOff As New PinListData
Private mpldVldInFqm2Standalone As New PinListData
Private mpldVldOutFqm2Standalone As New PinListData
Private mpldVldInFqm3Standalone As New PinListData
Private mpldVldOutFqm3Standalone As New PinListData
Private mpldIccLo1StaOff As New PinListData
Private mpldIccLo2StaOn As New PinListData
Private mpldIccLo1Path As New PinListData
Private mpldIccLoin1Off As New PinListData
Private mpldIccLoout1Off As New PinListData
Private mpldVldLoout1 As New PinListData
Private mpldVldFqm3 As New PinListData
Private mpldIccLoout2On As New PinListData
Private mpldIccLoin2On As New PinListData
Private mpldVldLoin2 As New PinListData
Private mpldVldInFqm3 As New PinListData
Private mpldVldLoout2 As New PinListData
Private mpldIccFqm3Off As New PinListData
Private mpldIccCasSplitOff As New PinListData
Private mpldIccCasCombOff As New PinListData
Private mpldIccFqm2Off As New PinListData
Private mslTstBufSweepValRead As New SiteLong
Private msdTstBufSweepValReadV As New SiteDouble
Private mpldaTxloIccResults(13) As New PinListData
Private mpldaTxloVldResults(8) As New PinListData

' ENUM REGISTER VARIABLES

'<AUTOCOMMENT:0.4:ENUM
'''@enum OptTxPldSel
'''This enumeration  Tx Pld Conf
'''@var OptTx1Pld_En
'''Index for Tx1 Pld Enable
'''@var OptTx1Pld_Disable
'''Index for Tx1 Pld Disable
'''@var OptTx2Pld_En
'''Index for Tx2 Pld Enable
'''@var OptTx2Pld_Disable
'''Index for Tx2 Pld Disable
'''@var OptTx3Pld_En
'''Index for Tx3 Pld Enable
'''@var OptTx3Pld_Disable
'''Index for Tx3 Pld Enable
'''@var OptTx4Pald_En
'''Index for Tx4 Pld Enable
'''@var OptTx4Pld_Disable
'''Index for Tx4 Pld Enable
'>AUTOCOMMENT:0.4:ENUM
Public Enum OptTxPldSel
    TxPldSelNone = 0
    TxPldSel15 = 1
    TxPldSel26 = 2
    TxPldSel37 = 3
    TxPldSel48 = 4
End Enum

'<AUTOCOMMENT:0.4:ENUM
'''@enum OptEnDis
'''This enumeration is generic for Enabling or Disabling bitfields
'''@var TxDisable
'''Index for Disabling
'''@var TxEnable
'''Index for Enabling
'>AUTOCOMMENT:0.4:ENUM
Public Enum OptEnDis
    TxDisable = 0
    TxEnable = -1
End Enum

'<AUTOCOMMENT:0.4:ENUM
'''@enum OptLoBias
'''This enumeration TX LO Bias configuration
'''@var OptLoBias_En
'''Index for Enable Tx Lo Bias
'''@var OptLoBias_Disable
'''Index for Disable Tx Lo Bias
'>AUTOCOMMENT:0.4:ENUM
Public Enum OptLoBias
    OptLoBias_En = 1 ' Enable Lo
    OptLoBias_Disable = 0 ' Disable Lo
End Enum

'<AUTOCOMMENT:0.4:ENUM
'''@enum OptLoTstBuf
'''This enumeration TX LOTSTBUF1/2 configuration
'''@var OptLoTstBuf_En
'''Index for Enable Tx LoTstBuf1/2
'''@var OptLoTstBuf_Disable
'''Index for Disable Tx LoTstBuf1/2
'>AUTOCOMMENT:0.4:ENUM
Private Enum OptLoTstBuf
    OptLoTstBuf1_En = 1
    OptLoTstBuf1_Disable = 0
    OptLoTstBuf2_En = 1
    OptLoTstBuf2_Disable = 0
End Enum

' MAIN TEST INSTANCES

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Test to measure the PLD voltage and trim that value to a target DC operating point
'''@details
'''The DC operating point must be chosen to make sure that it will not be 0V when PA is at maximum power
''' - @JAMA{16837086,[[PRTESPEC_TX]] TXGE: TXRF PLD Trimming}
'''
'''@param[in,out] psRelaySetup Relay setting
'''@param[in,out] Validating_ during program validation
'''@return [As Long] No value to be processed
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'''@log{2025-08-06, Mannsfeld Mario, Reset connections and bleeder-resistor\, @JIRA{RSIPPTE-255}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeTxrfTrimming( _
    psRelaySetup As String, _
    Optional Validating_ As Boolean = False) As Long
    
    On Error GoTo ErrHandler
    Dim lpldaUntrimPldV(7) As New PinListData
    Dim lpldTrimPldVCh15 As New PinListData
    Dim lpldTrimPldVCh26 As New PinListData
    Dim lpldTrimPldVCh37 As New PinListData
    Dim lpldTrimPldVCh48 As New PinListData
    
    Call InitInstance(Validating_)
    If Validating_ Then
        Exit Function
    End If
    
    ''' Initialize PinListData measurement variables
    Call TxgeTxrfPldInitPld(lpldaUntrimPldV())
    ''' Apply Levels and Timing
    Call DeviceSetup
    ''' Instrument setup for DCVI single-ended measurements on AMUX1 and AMUX2
    ''' - High impedance mode, meter voltage, 32 HW averages with sampling at 100kHz, bleeder resistors OFF
    Call TxgePldTrimmingDcviSetup("AMUX1_P_DCVI, AMUX2_P_DCVI")
 
    ''' Measurement and Trimming section with instrument reset and datalogging
    ''' -# Start the pattern
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName)
    ''' -# Disable the DFX protection
    Call Ctrx.DisableDfxGateProtection
    ''' -# Enable TX central biasing
    Call TxCentralBias(OptTxCentralBias_En)
    ''' -# Enable the local biasing for all TXRF channels
    Call TxgeTxrfBiasEnable
    ''' -# Measure untrimmed PLD voltage and perform PLD Trimming for Ch1 to Ch8
    Call TxgeTxrfPldMeasTrim(lpldaUntrimPldV(), lpldTrimPldVCh15, lpldTrimPldVCh26, lpldTrimPldVCh37, lpldTrimPldVCh48)
    ''' -# Set the Pld Trimming register in the EFuse cache
    Call TxgeTxrfPldTrimmValueToEfuseCache(lpldTrimPldVCh15, lpldTrimPldVCh26, lpldTrimPldVCh37, lpldTrimPldVCh48)
    ''' -# Enable back the DFX protection
    Call Ctrx.EnableDfxGateProtection
    ''' -# Soft-Reset and pattern stop
    Call Ctrx.SoftReset(peUpdateMethod:=ForceUpdate, pbAlarmToggleRequired:=True)
    Call Ctrx.DoppStopPattern
    
    ''' -# Reset the DCVI Instruments
    Call IfxHdw.Dcvi.Reset("AMUX1_P_DCVI, AMUX2_P_DCVI", peOption:=tlResetConnections + tlResetSettings, _
        pbResetBleederResistor:=True)
    
    ''' -# Datalogging for all measured values
    Call TxgeTxrfPldTrimlog(lpldaUntrimPldV(), lpldTrimPldVCh15, lpldTrimPldVCh26, lpldTrimPldVCh37, lpldTrimPldVCh48)
    
    tfTxgeTxrfTrimming = EndInstance
    Exit Function
ErrHandler:

    If AbortTest Then HardTesterReset Else Resume Next
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Txge TXLO LOIN & LOOUT Trimming
'''@details
''' Measure and Trim the values to be programmed in the TXLO bitfields
''' - @JAMA{22659319,[[PRTESPEC_TX]] TXGE: TXLO PLD VLD Trimming}
'''
'''@param[in] psRelaySetup Relay setting
'''@param[in] Validating_ during program validation
'''@return [As Long] No value to be processed.
'''
'''@log{2023-12-12, Olia Svyrydova, @JIRA{CTRXTE-3463}: TXLO trimming}
'''@log{2024-05-13, Wandji Lionel Wilfried, Update of some common FW-commands\, @JIRA{CTRXTE-4403,CTRXTE-4376}}
'''@log{2024-07-29, Neul Roland, @JIRA{CTRXTE-3821} force trim register update after soft reset}
'''@log{2024-11-12, Neul Roland, move SPI/JTAG support to library\, @JIRA{CTRXTE-5202}}
'''@log{2025-07-02, Channon Andrew, Revised common routines to adapt to 8188\, @JIRA{RSIPPTE-107}}
'''@log{2025-07-14, Channon Andrew, Revised to adapt to 8188\, @JIRA{RSIPPTE-157}}
'''@log{2025-08-06, Mannsfeld Mario, Reset connections and bleeder-resistor\, @JIRA{RSIPPTE-255}}
'''@log{2025-10-07, Holliday Dave, Renamed TxCtrlLoBias to call shared routine\, @JIRA{RSIPPTE-113}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeTxloTrimming( _
    psRelaySetup As String, _
    Optional Validating_ As Boolean = False) As Long
      
    Dim lpldaUntrimVldPldV(2) As New PinListData
    Dim lpldTrimPldV As New PinListData
    Dim lpldTrimVldV1 As New PinListData
    Dim lpldTrimVldV2 As New PinListData
        
    On Error GoTo ErrHandler
    
    Call InitInstance(Validating_)
    If Validating_ Then
        Exit Function
    End If
    
    ''' Initialize PinListData measurement variables
    Call TxgeTxloInitPld(lpldaUntrimVldPldV())
    ''' Apply Levels and Timing
    Call DeviceSetup
    ''' Instrument setup for DCVI single-ended measurements on AMUX1
    ''' - High impedance mode, meter voltage, 32 HW averages with sampling at 100kHz, bleeder resistors OFF
    ''' - Get the settling wait time for DCVI measurement.
    Call TxgePldTrimmingDcviSetup("AMUX1_P_DCVI")
    mdDcviWait = IfxHdw.Dcvi.GetStrobeWait("AMUX1_P_DCVI")
 
    ''' Measurement and Trimming section with instrument reset and datalogging
    ''' -# Start the pattern
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName)
    ''' -# Disable the DFX protection
    Call Ctrx.DisableDfxGateProtection
    ''' -# Enable TX central biasing
    Call TxCentralBias(OptTxCentralBias_En)
    ''' -# Enable the central TXLO biasing
    Call TxCtrlLoBias(OptLoBias_En)
    ''' -# Measure untrimmed VLD/PLD voltages and perform Txlo Vld/Pld trimming
    Call TxgeTxloPldVldMeasTrim(lpldaUntrimVldPldV(), lpldTrimVldV1, lpldTrimVldV2, lpldTrimPldV)
    ''' -# Set the Vld and Pld bitfields in the Trimming register in the EFuse cache
    Call TxgeTxloTrimmValueToEfuseCache(lpldTrimVldV1, lpldTrimPldV)
    ''' -# Enable back the DFX protection
    Call Ctrx.EnableDfxGateProtection
    ''' -# Soft-Reset and pattern stop
    Call Ctrx.SoftReset(peUpdateMethod:=ForceUpdate, pbAlarmToggleRequired:=True)
    Call Ctrx.DoppStopPattern
 
    ''' -# Reset the DCVI instruments
    Call IfxHdw.Dcvi.Reset("AMUX1_P_DCVI", peOption:=tlResetConnections + tlResetSettings, _
        pbResetBleederResistor:=True)
    
    ''' -# Datalogging for all measured values
    Call TxgeTxloTrimlog(lpldaUntrimVldPldV(), lpldTrimVldV1, lpldTrimVldV2, lpldTrimPldV)
    
    tfTxgeTxloTrimming = EndInstance
    Exit Function
    
ErrHandler:
    If AbortTest Then HardTesterReset Else Resume Next
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Test function to measure the current consumption of each block in TXLO and TXRF
''' - @JAMA{22282055,[[PRTESPEC_TX]] TXGE: TX DC levels and current consumption}
'''@details
''' Each block component is progressively enabled and the current consumption is monitored.
'''
'''@param[in] psRelaySetup Relay setting
'''@return [As Long] No value to be processed.
'''
'''@log{2025-09-10, Channon Andrew, Ported from 8191 and adapted for 8188A\, @JIRA{RSIPPTE-109}}
'''@log{2025-10-28, Channon Andrew, Remove Div8 tests\, @JIRA{RSIPPTE-411}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeCurrCons( _
    psRelaySetup As String, _
    Optional Validating_ As Boolean = False) As Long

    On Error GoTo ErrHandler
    Dim lwVddPre As New DSPWave
    Dim lwVddCorr As New DSPWave
    Dim lwVddPost As New DSPWave
    
    Call InitInstance(Validating_)
    If Validating_ Then
        Exit Function
    End If
    
    '''-# Initialization
    Call TxgeCurConsVariablesInit
    ''' Apply Levels and Timing
    Call DeviceSetup
    ''' Instrument setup for DIFFMETER on AMUX1, also DCVI for current measurements on VDD1V0TX
    Call TxAmuxDiffmeterSetup(DIFF_METER, tlDCDiffMeterModeHighSpeed, DiffMeterRange3p5V)
    ''' Get the wait time for the Diff setup measurement.
    mdPaDiffMeterWait = IfxHdw.DCDiffMeter.GetStrobeWait("AMUX1P_DIFF") * 2
    
    ''' Run pattern for preconditioning and measure uncompensated voltage supplies
    Call TxgeCurConsPreconditioning
    '''-# Read measured values from Preconditioning pattern and calculate Vdd compensation values
    Call TxgeCurConsReadVddPre
    '''-# Program the DCVI instruments for current measurements and apply compensation for MAX supply levels
    Call TxgeCurConsInstrSetup
    '''-# Measure the compensated supply voltages and Icc supply currents for Max then Min corners
    Call TxgeCurConsMeasCorners
    '''-# Measure the DPLL 1V2 Bias voltage
    Call TxgeCurConsDpllBias1V2
    '''-# Measure the DC currents using the AMUX pins
    Call TxgeCurConsDcCurrMeas
    '''-# The reset pattern ends with the necessary DUT deconditioning
    Call TxgeCurConsReset
    
    '''-# Calculate delta current values
    Call TxgeCurConsCalc
    ''' Datalogging
    Call TxgeCurConsDatalog

    ''' Instrument reset
    Call TxgeCurConsSupplyResetConfig
    
    tfTxgeCurrCons = EndInstance
    Exit Function
ErrHandler:
    If AbortTest Then HardTesterReset Else Resume Next
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Test function to measure the current consumption of each block in TXLO
''' - @JAMA{24143770,[[PRTESPEC_TX]] TXGE TXLO: DC levels and current consumption}
'''@details
''' Each block component is progressively enabled and the current consumption is monitored.
'''
'''@param[in] psRelaySetup Relay setting
'''@return [As Long] No value to be processed.
'''
'''@log{2025-11-11, Channon Andrew, initial implementation for 8188A\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeTxloCurrCons( _
    psRelaySetup As String, _
    Optional Validating_ As Boolean = False) As Long

    On Error GoTo ErrHandler
    Dim lwVddPre As New DSPWave
    Dim lwVddCorr As New DSPWave
    Dim lwVddPost As New DSPWave
    
    Call InitInstance(Validating_)
    If Validating_ Then
        Exit Function
    End If
    
    '''-# Initialization
    Call TxgeTxloCurConsVariablesInit
    ''' Apply Levels and Timing
    Call DeviceSetup
    ''' Instrument setup for DIFFMETER on AMUX1P
    Call TxAmuxDiffmeterSetup(DIFF_METER1, tlDCDiffMeterModeHighAccuracy, DiffMeterRange3p5V, _
        pbDcviVdd1V0TxMeasCurrSetup:=False)
    ''' Get the wait time for the Diff setup measurement.
    mdPaDiffMeterWait = IfxHdw.DCDiffMeter.GetStrobeWait("AMUX1P_DIFF")
    
    ''' Run pattern for VLD offset measurements
    Call TxgeTxloCurConsVldOffsets
    ''' Run pattern for preconditioning and measure uncompensated voltage supplies
    Call TxgeTxloCurConsPreconditioning
    '''-# Program the DCVI instruments for current measurements
    Call TxgeTxloCurConsInstrSetup
    '''-# Run pattern to measure the Icc supply currents and VLD voltages
    Call TxgeTxloCurConsMeas
    
    '''-# Calculate delta current values
    Call TxgeTxloCurConsCalc
    ''' Datalogging
    Call TxgeTxloCurConsDatalog

    ''' Instrument reset
    Call TxgeCurConsSupplyResetConfig
    
    tfTxgeTxloCurrCons = EndInstance
    Exit Function
ErrHandler:
    If AbortTest Then HardTesterReset Else Resume Next
End Function

' SUB-PROCEDURES OF MAIN

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Enable TX Local Bias for all 8 TX channels
'''@details
'''This routine should be called inside a DOPP block for pattern communication
'''
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'''@log{2025-09-01, Wandji Lionel Wilfried, Update on the tempco bitfield correction\, @JIRA{RSIPPTE-326}}
'''@log{2025-09-11, Neul Roland, replace hardcoded tempco with constant\, @JIRA{RSIPPTE-327}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfBiasEnable()

    Dim liTxChannelIx As Long
    
    ''' Enable the local biasing for all TXRF channels using a loop
    For liTxChannelIx = 1 To 8
        With CallByName(IfxRegMap, "NewRegTxTxrf" & liTxChannelIx & "DigCf", VbMethod)
            .bfPaDigInit = 1
            .bfPsDigInit = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & liTxChannelIx & "DigCf", VbMethod)
            .bfPaDigInit = 0
            .bfPsDigInit = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & liTxChannelIx & "PaCf", VbMethod)
            .bfEnIdac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & liTxChannelIx & "BiasCf", VbMethod)
            .bfEn = 1
            .bfTc = SR_TXRF_PA_BIAS_TC_INIT_DEF
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        
        Call Dopp.Wait(SR_TXRF_BIAS_CONF_EN_WAIT)
    Next liTxChannelIx
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief TxgeConfigTxAmuxPld for selecting Txrf PLD Forward Voltage channel connections from the local Amux
'''@details
'''Ensures all channels are in a reset state before selecting the wanted channels
'''
'''@param[in] peTxPldSel Txrf PLD Forward Voltage channel selection
'''@param[in] peEnable Select option to Enable or Disable the AmuxCtrl
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfPldAmuxSetup( _
    ByVal peTxPldSel As OptTxPldSel, _
    ByVal peEnable As OptEnDis)

    Dim lsTxPldA As String
    Dim lsTxPldB As String
    
    If peTxPldSel = TxPldSelNone Then
        Exit Sub
    End If
    
    lsTxPldA = CStr(peTxPldSel)
    lsTxPldB = CStr(peTxPldSel + 4)
    
    If Not peEnable Then
        ''' If not selected, reset the Txrf Channel local Amux
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldA & "SenseCf", VbMethod)
            .bfAmuxCtrl = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldB & "SenseCf", VbMethod)
            .bfAmuxCtrl = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
    Else
        ''' If selected, connect the corresponding TXRF# PLD forward voltage channel from the TXRF local AMUX
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldA & "SenseCf", VbMethod)
            .bfAmuxCtrl = &HA
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldB & "SenseCf", VbMethod)
            .bfAmuxCtrl = &HA
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Enable or Disable selected pairs of Txrf PLD channels
'''@details
'''Channel selections are paired: Ch1-Ch5, Ch2-Ch6, Ch3-Ch7, Ch4-Ch8
'''
'''@param[in] peTxPldSel Txrf PLD Forward Voltage channel selection
'''@param[in] peEnable Select option to Enable or Disable
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfPldConfig( _
    ByVal peTxPldSel As OptTxPldSel, _
    ByVal peEnable As OptEnDis)

    Dim lsTxPldA As String
    Dim lsTxPldB As String
    
    If peTxPldSel = TxPldSelNone Then
        Exit Sub
    End If
    
    lsTxPldA = CStr(peTxPldSel)
    lsTxPldB = CStr(peTxPldSel + 4)
    
    If Not peEnable Then
    ''' If Disable is selected then reset the bitfields in the PldCf register for the selected Txrf channels
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldA & "PldCf", VbMethod)
            .bfEn = 0
            .bfEnSourcefol = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldB & "PldCf", VbMethod)
            .bfEn = 0
            .bfEnSourcefol = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
    Else
    ''' If Enable is selected then set the EN bitfield in the PldCf register for the selected Txrf channels
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldA & "PldCf", VbMethod)
            .bfEn = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    
        With CallByName(IfxRegMap, "NewRegTxTxrf" & lsTxPldB & "PldCf", VbMethod)
            .bfEn = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
   
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Function initializes Plds
'''
'''@param[in,out] ppldaUntrimPldV PLD untrimmed voltage
'''
'''@log{2025-07-02, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfPldInitPld( _
    ByRef ppldaUntrimPldV() As PinListData)
    
    Dim li As Long
    For li = 0 To 3
        Set ppldaUntrimPldV(li) = IfxMath.Calc.PinListData.Create(-99, "AMUX2_P_DCVI")
    Next li
    For li = 4 To 7
        Set ppldaUntrimPldV(li) = IfxMath.Calc.PinListData.Create(-99, "AMUX1_P_DCVI")
    Next li

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief It implements the measurement sweep and trims the TXRF PLD voltage for Ch1 to Ch8
'''@details
'''Firstly measure the PLD voltage for all trimming codes and then trim the PLD voltage to be 0.53V
'''
'''@param[in,out] ppldaUntrimPldV Untrimmed PLD Voltage of all Chs
'''@param[in,out] ppldTrimPldVCh15 Trimmed PLD voltage of Ch1 and Ch5
'''@param[in,out] ppldTrimPldVCh26 Trimmed PLD voltage of Ch2 and Ch6
'''@param[in,out] ppldTrimPldVCh37 Trimmed PLD voltage of Ch3 and Ch7
'''@param[in,out] ppldTrimPldVCh48 Trimmed PLD voltage of Ch4 and Ch8
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'''@log{2025-07-31, Neul Roland, bringup fix\, @JIRA{RSIPPTE-215}}
'''@log{2025-08-11, Ibrahim Osama, remove the routing to Sadc as it is already handled
'''in ConfigGlobalAmux\, @JIRA{RSIPPTE-219}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfPldMeasTrim( _
    ByRef ppldaUntrimPldV() As PinListData, _
    ByRef ppldTrimPldVCh15 As PinListData, _
    ByRef ppldTrimPldVCh26 As PinListData, _
    ByRef ppldTrimPldVCh37 As PinListData, _
    ByRef ppldTrimPldVCh48 As PinListData)
    
    Dim lpldaCh15() As PinListData
    Dim lpldaCh26() As PinListData
    Dim lpldaCh37() As PinListData
    Dim lpldaCh48() As PinListData
    
    ''' Global Amux configured to connect PAD 1 and PAD 2 Differential to keep the PLD loads symmetric
    Call TxConfigGlobalAmux(OptAmuxPADConfType_Pad1En, OptAmuxConfType_Differential, _
                            OptAmuxPADConfType_Pad2En, OptAmuxConfType_Differential)
    ''' Configure Tx1 and Tx5 for measurement and trimming
    ''' -# Select the TXRF PLD forward voltage channel from the TXRF local AMUX
    Call TxgeTxrfPldAmuxSetup(TxPldSel15, TxEnable)
    ''' -# Enable the TXRF PLD
    Call TxgeTxrfPldConfig(TxPldSel15, TxEnable)
    ''' -# Wait for settling, measure and trim the TXRF PLD voltages returning the trimmed voltage and bitfield setting
    Call Dopp.Wait(WAIT_1MS)
    ppldTrimPldVCh15 = TxgePldTrimm(TXPLD_CH15, lpldaCh15())
    ''' -# Disable the TXRF PLD
    Call TxgeTxrfPldConfig(TxPldSel15, TxDisable)
    ''' -# Reset the TXRF local AMUX
    Call TxgeTxrfPldAmuxSetup(TxPldSel15, TxDisable)
    
    ''' Configure Tx2 and Tx6 for measurement and trimming by repeating the steps for Tx1 and Tx5 above
    Call TxgeTxrfPldAmuxSetup(TxPldSel26, TxEnable)
    Call TxgeTxrfPldConfig(TxPldSel26, TxEnable)
    Call Dopp.Wait(WAIT_1MS)
    ppldTrimPldVCh26 = TxgePldTrimm(TXPLD_CH26, lpldaCh26())
    Call TxgeTxrfPldConfig(TxPldSel26, TxDisable)
    Call TxgeTxrfPldAmuxSetup(TxPldSel26, TxDisable)
    
    ''' Configure Tx3 and Tx7 for measurement and trimming by repeating the steps for Tx1 and Tx5 above
    Call TxgeTxrfPldAmuxSetup(TxPldSel37, TxEnable)
    Call TxgeTxrfPldConfig(TxPldSel37, TxEnable)
    Call Dopp.Wait(WAIT_1MS)
    ppldTrimPldVCh37 = TxgePldTrimm(TXPLD_CH37, lpldaCh37())
    Call TxgeTxrfPldConfig(TxPldSel37, TxDisable)
    Call TxgeTxrfPldAmuxSetup(TxPldSel37, TxDisable)
    
    ''' Configure Tx4 and Tx8 for measurement and trimming by repeating the steps for Tx1 and Tx5 above
    Call TxgeTxrfPldAmuxSetup(TxPldSel48, TxEnable)
    Call TxgeTxrfPldConfig(TxPldSel48, TxEnable)
    Call Dopp.Wait(WAIT_1MS)
    ppldTrimPldVCh48 = TxgePldTrimm(TXPLD_CH48, lpldaCh48())
    Call TxgeTxrfPldConfig(TxPldSel48, TxDisable)
    Call TxgeTxrfPldAmuxSetup(TxPldSel48, TxDisable)

     ''' Return the untrimmed PLD voltage values from all trim sweeps measured at bitfield default settings (PLD=3)
    For Each Site In TheExec.Sites
        ppldaUntrimPldV(0) = lpldaCh15(3).Pins("AMUX2_P_DCVI").Value
        ppldaUntrimPldV(1) = lpldaCh26(3).Pins("AMUX2_P_DCVI").Value
        ppldaUntrimPldV(2) = lpldaCh37(3).Pins("AMUX2_P_DCVI").Value
        ppldaUntrimPldV(3) = lpldaCh48(3).Pins("AMUX2_P_DCVI").Value
        ppldaUntrimPldV(4) = lpldaCh15(3).Pins("AMUX1_P_DCVI").Value
        ppldaUntrimPldV(5) = lpldaCh26(3).Pins("AMUX1_P_DCVI").Value
        ppldaUntrimPldV(6) = lpldaCh37(3).Pins("AMUX1_P_DCVI").Value
        ppldaUntrimPldV(7) = lpldaCh48(3).Pins("AMUX1_P_DCVI").Value
    Next Site
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Trimming procedure for TXRF PLD Trimming
'''@details
'''Two TXRF PLD channels are measured together. Channels 1-4 are measured using AMUX2 and channels 5-8 are measured
''' using AMUX1. The trim values are calculated from the sweep and these are added to the measured values and returned
''' in the single PinListData object.
''' The pairings used are Ch1-Ch5, Ch2-Ch6, Ch3-Ch7, Ch4-Ch8
'''
'''@param[in] piChSelec Channel select
'''@param[in,out] ppldaRawData contains LDO Voltage measurement values corresponding to all trimming codes
'''@return [As PinListData] contains measured voltages from AMUX1 and AMUX2 and trim values for 1 pair of TX channels
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function TxgePldTrimm( _
    ByVal piChSelec As Long, _
    ByRef ppldaRawData() As PinListData) As PinListData
    
    Dim lbBigJumpsAllowed As Boolean
    Dim ldTxStartVal As Double
    Dim ldTxEndVal As Double
    Dim ldTxDefaultVal As Double
    Dim ldTxTargetVal As Double
    Dim ldTxAvgCntVal As Double
    Dim lbTxDecending As Boolean
    Dim ldWaitSettle As Double
    Dim lsMeasPin As String
    Dim loTx1TrimReg As RegTxTxrf1PldCf
    Dim loTx2TrimReg As RegTxTxrf2PldCf
    Dim loTx3TrimReg As RegTxTxrf3PldCf
    Dim loTx4TrimReg As RegTxTxrf4PldCf
    Dim loTx5TrimReg As RegTxTxrf5PldCf
    Dim loTx6TrimReg As RegTxTxrf6PldCf
    Dim loTx7TrimReg As RegTxTxrf7PldCf
    Dim loTx8TrimReg As RegTxTxrf8PldCf
    Dim loBfPldConf1 As Bitfield
    Dim loBfPldConf2 As Bitfield
    Dim loBfPldConf3 As Bitfield
    Dim loBfPldConf4 As Bitfield
    
    lsMeasPin = "AMUX1_P_DCVI,AMUX2_P_DCVI"
    ''' -# Trimming register selection
    If (piChSelec = TXPLD_CH15) Then
        Set loTx1TrimReg = IfxRegMap.NewRegTxTxrf1PldCf.withBfEn(1)
        Set loBfPldConf1 = loTx1TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf3 = loTx1TrimReg.bfCtrlRevGetBitField
        Set loTx5TrimReg = IfxRegMap.NewRegTxTxrf5PldCf.withBfEn(1)
        Set loBfPldConf2 = loTx5TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf4 = loTx5TrimReg.bfCtrlRevGetBitField
    ElseIf (piChSelec = TXPLD_CH26) Then
        Set loTx2TrimReg = IfxRegMap.NewRegTxTxrf2PldCf.withBfEn(1)
        Set loBfPldConf1 = loTx2TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf3 = loTx2TrimReg.bfCtrlRevGetBitField
        Set loTx6TrimReg = IfxRegMap.NewRegTxTxrf6PldCf.withBfEn(1)
        Set loBfPldConf2 = loTx6TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf4 = loTx6TrimReg.bfCtrlRevGetBitField
    ElseIf (piChSelec = TXPLD_CH37) Then
        Set loTx3TrimReg = IfxRegMap.NewRegTxTxrf3PldCf.withBfEn(1)
        Set loBfPldConf1 = loTx3TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf3 = loTx3TrimReg.bfCtrlRevGetBitField
        Set loTx7TrimReg = IfxRegMap.NewRegTxTxrf7PldCf.withBfEn(1)
        Set loBfPldConf2 = loTx7TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf4 = loTx7TrimReg.bfCtrlRevGetBitField
    Else
        Set loTx4TrimReg = IfxRegMap.NewRegTxTxrf4PldCf.withBfEn(1)
        Set loBfPldConf1 = loTx4TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf3 = loTx4TrimReg.bfCtrlRevGetBitField
        Set loTx8TrimReg = IfxRegMap.NewRegTxTxrf8PldCf.withBfEn(1)
        Set loBfPldConf2 = loTx8TrimReg.bfCtrlFwdGetBitField
        Set loBfPldConf4 = loTx8TrimReg.bfCtrlRevGetBitField
    End If
    
    lbBigJumpsAllowed = PLD_BIG_JUMPS_ALLOWED
    ldTxStartVal = PLD_TX_START_VAL
    ldTxEndVal = PLD_TX_END_VAL
    ldTxDefaultVal = PLD_TX_DEFAULT_VAL
    ldTxTargetVal = PLD_TX_TARGET_VAL
    ldTxAvgCntVal = PLD_TX_AVG_CNT_VAL
    lbTxDecending = PLD_TX_DECENDING
    ldWaitSettle = TX_MEAS_SET_WAIT
        
    ''' -# Setup the Trimming Object for sweeping
    Call Ctrx.Trimming.SetupWithBf(Dcvi, ldTxStartVal, ldTxEndVal, ldTxDefaultVal, ldTxTargetVal, _
                             lsMeasPin, ldTxAvgCntVal, lbBigJumpsAllowed, ldWaitSettle, _
                             loBfPldConf1, loBfPldConf2, lbTxDecending, _
                             pbTakeClosesdHigherThen:=True, poBitField3:=loBfPldConf3, poBitField4:=loBfPldConf4)
    ''' -# Do the full sweep trimming
    Set TxgePldTrimm = Ctrx.Trimming.SweepWithRawData(ppldaRawData())
            
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function saves the trimming values to the Efuse cache
'''@details
'''TXRF PLD trimming values are saved to the Efuse Cache
'''
'''@param[in,out] ppldaTrimPldVCh15 Trimmed Pld values of Ch1&Ch5
'''@param[in,out] ppldaTrimPldVCh26 Trimmed Pld values of Ch2&Ch6
'''@param[in,out] ppldaTrimPldVCh37 Trimmed Pld values of Ch3&Ch7
'''@param[in,out] ppldaTrimPldVCh48 Trimmed Pld values of Ch4&Ch8
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfPldTrimmValueToEfuseCache( _
     ppldaTrimPldVCh15 As PinListData, _
     ppldaTrimPldVCh26 As PinListData, _
     ppldaTrimPldVCh37 As PinListData, _
     ppldaTrimPldVCh48 As PinListData)
    
    Dim lslaTrimPldV(7) As New SiteLong
    
    ''' -# Check for Trim errors in sites (if an error save the 0 value in the cache)
    lslaTrimPldV(0) = ppldaTrimPldVCh15.Pins("AMUX2_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh15.Pins("AMUX2_P_DCVI"))
    lslaTrimPldV(1) = ppldaTrimPldVCh26.Pins("AMUX2_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh26.Pins("AMUX2_P_DCVI"))
    lslaTrimPldV(2) = ppldaTrimPldVCh37.Pins("AMUX2_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh37.Pins("AMUX2_P_DCVI"))
    lslaTrimPldV(3) = ppldaTrimPldVCh48.Pins("AMUX2_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh48.Pins("AMUX2_P_DCVI"))
    lslaTrimPldV(4) = ppldaTrimPldVCh15.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh15.Pins("AMUX1_P_DCVI"))
    lslaTrimPldV(5) = ppldaTrimPldVCh26.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh26.Pins("AMUX1_P_DCVI"))
    lslaTrimPldV(6) = ppldaTrimPldVCh37.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh37.Pins("AMUX1_P_DCVI"))
    lslaTrimPldV(7) = ppldaTrimPldVCh48.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
            (TX_TRIM_ERR_0, ppldaTrimPldVCh48.Pins("AMUX1_P_DCVI"))
  
    ''' -# Save PLD Trim value to EFuse Cache via the register object that can be retrieved from the cache
    With Ctrx.EFuse.RegAppWord2
        .bfTxPldCtrl1ms = lslaTrimPldV(0)
        .bfTxPldCtrl2ms = lslaTrimPldV(1)
        .bfTxPldCtrl3ms = lslaTrimPldV(2)
        .bfTxPldCtrl4ms = lslaTrimPldV(3)
        .bfTxPldCtrl5ms = lslaTrimPldV(4)
        .bfTxPldCtrl6ms = lslaTrimPldV(5)
        .bfTxPldCtrl7ms = lslaTrimPldV(6)
        .bfTxPldCtrl8ms = lslaTrimPldV(7)
        Call Ctrx.EFuse.Cache.AppWord(2).SetValue(.Self.Value)
    End With

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Datalog measured PLD voltages for Ch1-Ch8 along with final trim settings
'''@details
'''The datalog includes default and trimmed PLD Voltages together with finalized Trim code
'''
'''@param[in,out] ppldaUntrimPldV is Untrimmed PLD Voltage of Ch1 to Ch8
'''@param[in,out] ppldTrimPldVCh15 is Trimmed PLD Value and Voltage of Ch1&Ch5
'''@param[in,out] ppldTrimPldVCh26 is Trimmed PLD Value and Voltage of Ch2&Ch6
'''@param[in,out] ppldTrimPldVCh37 is Trimmed PLD Value and Voltage of Ch3&Ch7
'''@param[in,out] ppldTrimPldVCh48 is Trimmed PLD Value and Voltage of Ch4&Ch8
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxrfPldTrimlog( _
    ppldaUntrimPldV() As PinListData, _
    ppldTrimPldVCh15 As PinListData, _
    ppldTrimPldVCh26 As PinListData, _
    ppldTrimPldVCh37 As PinListData, _
    ppldTrimPldVCh48 As PinListData)

    Dim li As Long
    ''' -# Datalog the measured untrimmed voltages at default bitfield settings for all channels
    For li = 0 To 7
        Call IfxEnv.Datalog(ppldaUntrimPldV(li))
    Next li

    ''' -# Datalog the trimmed bitfield settings
    Call IfxEnv.Datalog(ppldTrimPldVCh15.Pins("AMUX2_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh26.Pins("AMUX2_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh37.Pins("AMUX2_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh48.Pins("AMUX2_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh15.Pins("AMUX1_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh26.Pins("AMUX1_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh37.Pins("AMUX1_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldVCh48.Pins("AMUX1_P_DCVI"))
    ''' -# Datalog the measured trimmed voltages
    Call IfxEnv.Datalog(ppldTrimPldVCh15.Pins("AMUX2_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh26.Pins("AMUX2_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh37.Pins("AMUX2_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh48.Pins("AMUX2_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh15.Pins("AMUX1_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh26.Pins("AMUX1_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh37.Pins("AMUX1_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldVCh48.Pins("AMUX1_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Configure AMUX pins for TXRF and TXLO PLD voltage measurements for Trimming
'''@details
'''This is a general instrumentation setup which should be done after the device setup and before measurement
'''
'''@param[in] psPins Specify the DCVI pins to be set up. Can be more than 1 pin.
'''
'''@log{2025-07-18, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-107}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgePldTrimmingDcviSetup( _
    ByVal psPins As String)
    
    ''' Instrument setup for DCVI single-ended measurements on selected AMUX pins
    ''' - High impedance mode, meter voltage, 32 HW averages with sampling at 100kHz, bleeder resistors OFF
    Call IfxHdw.Dcvi.Setup(psPins, tlDCVIModeHighImpedance, DCVI_DEFAULT_VOLTAGE, DCVI_DEFAULT_CURRENT, _
        DcviRange1A, , tlDCVIMeterVoltage, HW_AVERAGE_32, peBleederSetting:=tlDCVIBleederResistorOff)
   
    TheHdw.Dcvi.Pins(psPins).Capture.SampleRate = 100 * kHz
End Sub



'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Routine to measure and trim the TXLO PLD and VLD voltages
'''@details As there is only 1 control bitfield for the LOIN VLD, trimming is only done for TXLO LOIN 1 VLD.
''' The trimmed bitfield is applied for the TXLO LOIN 2 VLD during the second voltage measurement.
'''
'''@param[in,out] ppldaUntrimVldV is Untrimmed VLD Voltage of LoIn1 and LoIn2
'''@param[in,out] ppldTrimVldV1 is Trimmed VLD Value and Voltage of LoIn1
'''@param[in,out] ppldTrimVldV2 is Trimmed VLD Voltage of LoIn2
'''@param[in,out] ppldTrimPldV is Trimmed PLD Value and Voltage of LoOut
'''
'''@log{2023-12-19, Olia Svyrydova, Initial Version}
'''@log{2023-12-12, Olia vommi, @JIRA{CTRXTE-3838}: Name changing}
'''@log{2024-11-12, Neul Roland, move SPI/JTAG support to library\, @JIRA{CTRXTE-5202}}
'''@log{2025-01-22, Andrew Channon, Code cleanup\, @JIRA{CTRXTE-5531}}
'''@log{2025-07-02, Channon Andrew, Renamed TxGamuxConfig to adapt to 8188\, @JIRA{RSIPPTE-107}}
'''@log{2025-07-02, Channon Andrew, Re-worked to adapt to 8188\, @JIRA{RSIPPTE-157}}
'''@log{2025-08-15, Wandji Lionel Wilfried, LOIN1 VLD untrimmed measurement incorrect\, @JIRA{RSIPPTE-272}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloPldVldMeasTrim( _
    ByRef ppldaUntrimVldPldV() As PinListData, _
    ByRef ppldTrimVldV1 As PinListData, _
    ByRef ppldTrimVldV2 As PinListData, _
    ByRef ppldTrimPldV As PinListData)
    
    Dim lpldaLoIn1() As PinListData
    Dim lpldaLoOut() As PinListData
    Dim lslTrimVldLoIn1 As New SiteLong

    ''' Configure the Global Amux to connect PAD 1 differential and disable PAD 2 of TX Mux
    Call TxConfigGlobalAmux(OptAmuxPADConfType_Pad1En, OptAmuxConfType_Differential, _
                        OptAmuxPADConfType_Pad2Disable, OptAmuxConfType_Disconnect)
    ''' Measurement and Trimming section:
    ''' -# Select the LOIN1 input VLD channel from the TXLO local AMUX
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x14
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' -# Enable the TXLO LOIN1
    Call TxLoLoIn12(OptLoIn1_En, OptLoIn2_Disable)
    ''' -# Wait for settling, measure the untrimmed voltage on LOIN1
    Call Dopp.Wait(WAIT_1MS)
    Call Dopp.Dcvi("AMUX1_P_DCVI").Meter.Strobe(mdDcviWait)
    Call Ctrx.DoppPausePattern
    ppldaUntrimVldPldV(0) = TheHdw.Dcvi("AMUX1_P_DCVI").Meter.Read(tlNoStrobe)
    ''' -# Measure and trim the VLD on LOIN1 returning the trimmed voltage and bitfield setting
    ppldTrimVldV1 = TxgeTxloTrimm(LOIN, lpldaLoIn1())
    ''' -# Disable the TXLO LOIN1
    Call TxLoLoIn12(OptLoIn1_Disable, OptLoIn2_Disable)
    
    ''' -# Select the LOIN2 input VLD channel from the TXLO local AMUX
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x15
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' -# Enable the TXLO LOIN2
    Call TxLoLoIn12(OptLoIn1_Disable, OptLoIn2_En)
    ''' -# Wait for settling, measure the untrimmed voltage on LOIN2
    Call Dopp.Wait(WAIT_1MS)
    Call Dopp.Dcvi("AMUX1_P_DCVI").Meter.Strobe(mdDcviWait)
    Call Ctrx.DoppPausePattern
    ppldaUntrimVldPldV(1) = TheHdw.Dcvi("AMUX1_P_DCVI").Meter.Read(tlNoStrobe)
    ''' -# Write the VLD trimm value found from LOIN1 into LOIN2 Ctrl register
    lslTrimVldLoIn1 = ppldTrimVldV1.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
        (TX_TRIM_ERR_0, ppldTrimVldV1.Pins("AMUX1_P_DCVI"))
    With IfxRegMap.NewRegTxTxloLoin2Ctrl
        .bfCtrlLoin2IdacVldms = lslTrimVldLoIn1
        Call Ctrx.WriteVolatileRegister(.Self)
    End With
    ''' -# Wait for settling, measure the trimmed voltage on LOIN2
    Call Dopp.Wait(WAIT_1MS)
    Call Dopp.Dcvi("AMUX1_P_DCVI").Meter.Strobe(mdDcviWait)
    Call Ctrx.DoppPausePattern
    ppldTrimVldV2 = TheHdw.Dcvi("AMUX1_P_DCVI").Meter.Read(tlNoStrobe)
    ''' -# Disable the TXLO LOIN2
    Call TxLoLoIn12(OptLoIn1_Disable, OptLoIn2_Disable)
    
    ''' -# Select the LOOUT input PLD channel from the TXLO local AMUX
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x19
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' -# Enable the TXLO LOOUT2
    Call TxLoLoOut(OptLoOut1_Disable, OptLoOut2_En)
    ''' -# Wait for settling, measure the untrimmed voltage on LOOUT
    Call Dopp.Wait(WAIT_1MS)
    Call Dopp.Dcvi("AMUX1_P_DCVI").Meter.Strobe(mdDcviWait)
    Call Ctrx.DoppPausePattern
    ppldaUntrimVldPldV(2) = TheHdw.Dcvi("AMUX1_P_DCVI").Meter.Read(tlNoStrobe)
    ''' -# Measure and trim the PLD on LOOUT2 returning the trimmed voltage and bitfield setting
    ppldTrimPldV = TxgeTxloTrimm(LOOUT, lpldaLoOut())
    ''' -# Disable the TXLO LOOUT2
    Call TxLoLoOut(OptLoOut1_Disable, OptLoOut2_Disable)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Trimming procedure for TXLO Trimming
'''
'''@param[in] piChSelec Channel select
'''@param[out] ppldaRawData contains LDO Voltage measurement values corresponding to all trimming codes
'''@Return[out] function return Trim value
'''
'''@param[in] piChSelec Select LOIN or LOOUT
'''@param[in] pdTxTargetVal Target Value
'''@Return[out] function return Trim value
'''
'''@log{2023-12-20, Olia Svyrydova, Initial Version}
'''@log{2025-07-15, Channon Andrew, Re-worked to adapt to 8188\, @JIRA{RSIPPTE-157}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function TxgeTxloTrimm( _
    ByVal piChSelec As Long, _
    ByRef ppldaRawData() As PinListData) As PinListData
    
    Dim lbBigJumpsAllowed As Boolean
    Dim ldTxStartVal As Double
    Dim ldTxEndVal As Double
    Dim ldTxDefaultVal As Double
    Dim ldTxTargetVal As Double
    Dim ldTxAvgCntVal As Double
    Dim lbTxDecending As Boolean
    Dim ldWaitSettle As Double
    Dim lsMeasPin As String
    Dim loTxLoinTrimReg As RegTxTxloLoin1Ctrl
    Dim loTxLooutTrimReg As RegTxTxloLoout2Ctrl
    Dim loBfTxloConf As Bitfield

    ''' -# Trimming Register selection
    If (piChSelec = LOIN) Then
        Set loTxLoinTrimReg = IfxRegMap.NewRegTxTxloLoin1Ctrl
        Set loBfTxloConf = loTxLoinTrimReg.bfCtrlLoin1IdacVldGetBitField
        ldTxEndVal = VLD_TX_END_VAL
        ldTxDefaultVal = VLD_TX_DEFAULT_VAL
        ldTxTargetVal = VLD_TXLO_TARGET_VAL
    ElseIf (piChSelec = LOOUT) Then
        Set loTxLooutTrimReg = IfxRegMap.NewRegTxTxloLoout2Ctrl
        Set loBfTxloConf = loTxLooutTrimReg.bfCtrlIdacPldGetBitField
        ldTxEndVal = PLD_TX_END_VAL
        ldTxDefaultVal = PLD_TX_DEFAULT_VAL
        ldTxTargetVal = PLD_TX_TARGET_VAL
    End If

    lbBigJumpsAllowed = PLD_BIG_JUMPS_ALLOWED
    ldTxStartVal = PLD_TX_START_VAL
    ldTxAvgCntVal = PLD_TX_AVG_CNT_VAL
    lbTxDecending = PLD_TX_DECENDING
    ldWaitSettle = WAIT_500US
    lsMeasPin = "AMUX1_P_DCVI"

    ''' -# Setup the Trimming Object for sweeping
    Call Ctrx.Trimming.SetupWithBf(Dcvi, _
        ldTxStartVal, _
        ldTxEndVal, _
        ldTxDefaultVal, _
        ldTxTargetVal, _
        lsMeasPin, _
        ldTxAvgCntVal, _
        lbBigJumpsAllowed, _
        ldWaitSettle, _
        loBfTxloConf, , _
        lbTxDecending, _
        pbTakeClosesdHigherThen:=True)
    ''' -# Do the full sweep trimming
    Set TxgeTxloTrimm = Ctrx.Trimming.SweepWithRawData(ppldaRawData())
            
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function save the trimming values to the Efuse cache
'''@details TXLO PLD and VLD Trimming values to Efuse Cache
'''
'''@param[in] ppldTrimLOIN Trimmed VLD Voltage for LOIN1
'''@param[in] ppldTrimLOOUT Trimmed PLD Voltage of LOUT
'''
'''@log{2023-12-21, Olia Svyrydova, Initial Version}
'''@log{2025-07-14, Channon Andrew, Revised to adapt to 8188\, @JIRA{RSIPPTE-157}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloTrimmValueToEfuseCache( _
    ppldTrimLOIN As PinListData, _
    ppldTrimLOOUT As PinListData)
     
    Dim lslTrimVldLoIn1 As New SiteLong
    Dim lslTrimPldLoout As New SiteLong
     
    ''' -# Check for Trim errors in sites (if an error save the 0 value in the cache)
    lslTrimVldLoIn1 = ppldTrimLOIN.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
        (TX_TRIM_ERR_0, ppldTrimLOIN.Pins("AMUX1_P_DCVI"))
    lslTrimPldLoout = ppldTrimLOOUT.Pins("AMUX1_P_DCVI").Compare(EqualTo, TRIMM_ERROR).If _
        (TX_TRIM_ERR_0, ppldTrimLOOUT.Pins("AMUX1_P_DCVI"))
  
    ''' -# Save VLD, PLD Trim value to EFuse Cache via the register object that can be retrieved from the cache
    ''' -# Also populate the TXLO SLOPE parameters in the EFuse Cache register object setting fixed default values
    With Ctrx.EFuse.RegAppWord11
        .bfTxloTc1Slope1V8 = TRIM_TXLO_TC1_SLOPE
        .bfTxloTc2Slope1V8 = TRIM_TXLO_TC2_SLOPE
        .bfTxloTc3Slope1V8 = TRIM_TXLO_TC3_SLOPE
        .bfTxloTc4Slope1V8 = TRIM_TXLO_TC4_SLOPE
        .bfTxloTc5Slope1V8 = TRIM_TXLO_TC5_SLOPE
        .bfTxloLoinCtrlVldin1V8ms = lslTrimVldLoIn1
        .bfTxloLooutCtrlPldout1V8ms = lslTrimPldLoout
        Call Ctrx.EFuse.Cache.AppWord(11).SetValue(.Self.Value)
    End With

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Datalog measured PLD and VLD voltages for LoIn1, LoIn2, LoOut along with final trim settings
'''@details
'''The datalog includes default and trimmed VLD and PLD Voltages together with finalized Trim code
'''
'''@param[in] ppldaUntrimVldPldV is Untrimmed VLD Voltage of LoIn1 and LoIn2
'''@param[in] ppldTrimVldV1 is Trimmed VLD Value and Voltage of LoIn1
'''@param[in] ppldTrimVldV2 is Trimmed VLD Voltage of LoIn2
'''@param[in] ppldTrimPldV is Trimmed PLD Value and Voltage of LoOut
'''
'''@log{2023-12-21, Olia Svyrydova, Initial Version}
'''@log{2025-07-07, Channon Andrew, Re-worked to adapt to 8188\, @JIRA{RSIPPTE-157}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloTrimlog( _
    ByRef ppldaUntrimVldPldV() As PinListData, _
    ByRef ppldTrimVldV1 As PinListData, _
    ByRef ppldTrimVldV2 As PinListData, _
    ByRef ppldTrimPldV As PinListData)
     
    ''' -# Datalog the untrimmed voltages for all VLD and PLD paths
    Call IfxEnv.Datalog(ppldaUntrimVldPldV(0))
    Call IfxEnv.Datalog(ppldaUntrimVldPldV(1))
    Call IfxEnv.Datalog(ppldaUntrimVldPldV(2))

    ''' -# Datalog the trimmed bitfield settings
    Call IfxEnv.Datalog(ppldTrimVldV1.Pins("AMUX1_P_DCVI"))
    Call IfxEnv.Datalog(ppldTrimPldV.Pins("AMUX1_P_DCVI"))
    
    ''' -# Datalog the measured trimmed voltages
    Call IfxEnv.Datalog(ppldTrimVldV1.Pins("AMUX1_P_DCVI_VAL"), , , VLD_TXLO_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimVldV2.Pins("AMUX1_P_DCVI"), , , VLD_TXLO_TARGET_VAL)
    Call IfxEnv.Datalog(ppldTrimPldV.Pins("AMUX1_P_DCVI_VAL"), , , PLD_TX_TARGET_VAL)
   
    ''' -# Datalog the TXLO Slope default values
    Call IfxEnv.Datalog(TRIM_TXLO_TC1_SLOPE)
    Call IfxEnv.Datalog(TRIM_TXLO_TC2_SLOPE)
    Call IfxEnv.Datalog(TRIM_TXLO_TC3_SLOPE)
    Call IfxEnv.Datalog(TRIM_TXLO_TC4_SLOPE)
    Call IfxEnv.Datalog(TRIM_TXLO_TC5_SLOPE)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Function initializes Plds
'''
'''@param[in,out] ppldaUntrimVldPldV VLD untrimmed voltage
'''
'''@log{2025-07-15, Channon Andrew, Initial Version}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloInitPld( _
    ByRef ppldaUntrimVldPldV() As PinListData)
    
    Dim li As Long
    For li = 0 To 2
        Set ppldaUntrimVldPldV(li) = IfxMath.Calc.PinListData.Create(-99, "AMUX1_P_DCVI")
    Next li

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function initializes the variables defined at module level
'''
'''@log{2025-09-11, Channon Andrew, Initial Version}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsVariablesInit()

    Dim lsPins As String
    Dim lsVdd1v8Pin As String
    Dim liCount As Long
        
    If IfxStd.TesterHw.IsFrontendHw Then
        lsVdd1v8Pin = "VDD1V8TX_DCVI"
    Else
        lsVdd1v8Pin = "VDD1V8_DCVI"
    End If
    
    msHwSetTpConfig = IfxRadar.TpConfig("HWSETUP")
    
    If (msHwSetTpConfig = "FE") Then
        lsPins = IfxMath.Util.String_.Join(",", lsVdd1v8Pin, "VDD1V0TX_MDCVI", "VDD1V8DPLL_DCVI")
        msPatHw = "FE"
    Else
        lsPins = IfxMath.Util.String_.Join(",", lsVdd1v8Pin, "VDD1V0TX_MDCVI")
        msPatHw = "BE"
    End If

    Set msdPreOpDtct = Nothing
    Set msdLowPowDtct = Nothing
    Set msdOpDtct = Nothing
    Set msdExitPreOpErr = Nothing
    Set msdGotoOpErr = Nothing
    Set moFwCmdGetTempPre = Nothing
    Set moFwCmdGetTempPost = Nothing
    
    Set mwVddCorrMin = New DSPWave
    Set mwVddCorrMax = New DSPWave
    Set msdIccVdd1V0MaxCornerSum = New SiteDouble
    Set msdIccVdd1V0MinCornerSum = New SiteDouble
    Set mpldaCurrResults(58) = New PinListData
    
    For liCount = 0 To 16
        mpldaIddMaxCorner(liCount) = IfxMath.Calc.PinListData.Create(-999, lsPins)
        mpldaIddMinCorner(liCount) = IfxMath.Calc.PinListData.Create(-999, lsPins)
    Next liCount
    For liCount = 0 To 15
        mpldaVddMaxCorner(liCount) = IfxMath.Calc.PinListData.Create(-999, lsPins)
        mpldaVddMinCorner(liCount) = IfxMath.Calc.PinListData.Create(-999, lsPins)
    Next liCount
    For liCount = 0 To 3
        mpldaTxBiasCurr(liCount) = IfxMath.Calc.PinListData.Create(-999, "AMUX1_P_DCVI,AMUX1_N_DCVI")
        mpldaTxrfIbgCurr(liCount) = IfxMath.Calc.PinListData.Create(-999, "AMUX2_P_DCVI,AMUX2_N_DCVI")
        mpldaTxrfIbgCurr(liCount + 4) = IfxMath.Calc.PinListData.Create(-999, "AMUX1_P_DCVI,AMUX1_N_DCVI")
        mpldaTxrfIbgCurr(liCount + 8) = IfxMath.Calc.PinListData.Create(-999, "AMUX2_P_DCVI,AMUX2_N_DCVI")
        mpldaTxrfIbgCurr(liCount + 12) = IfxMath.Calc.PinListData.Create(-999, "AMUX1_P_DCVI,AMUX1_N_DCVI")
    Next liCount
    For liCount = 0 To 5
        mpldaDpllBiasCurr(liCount) = IfxMath.Calc.PinListData.Create(-999, "AMUX2_P_DCVI")
    Next liCount
    For liCount = 0 To 4
        mpldaTxloBiasCurr(liCount) = IfxMath.Calc.PinListData.Create(-999, "AMUX1_P_DCVI,AMUX1_N_DCVI")
    Next liCount
    mpldDpllSense1V2 = IfxMath.Calc.PinListData.Create(-999, "AMUX1P_DIFF")
    msPinsCurrCons = lsPins
    msVdd1v8Tx = lsVdd1v8Pin

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Run the Dopp pattern to measure the Vdd supplies uncompensated.
'''@details Measures the Vdd voltage for VDD0V9RF, VDD0V9PA and VDD1V8PA DCVIs using the AMUX. For the VddPA values,
''' one Tx channel is enabled, default TXCHANNEL1.
''' In total, 11 voltage measurements are made, all using AMUX1P_DIFF.
'''@param[out] poFwCmdExitPreOperation result from PreoptoOperation passed out for datalogging
'''@param[out] poFwCmdGotoOperation
'''@log{2025-02-19, Channon Andrew, initial implementation @JIRA{CTRXTE-5553} }
'''@log{2025-06-19, Channon Andrew, update parameter name to OptTxChSel_NoTx @JIRA{CTRXTE-5879} }
'''@log{2025-09-10, Channon Andrew, Ported from 8191 and adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsPreconditioning()
    
    Dim lsPatNameExt As String
    Dim liTxChannel As Long
    Dim loPreOpCapt As FirmwareVariable
    Dim loLowPowCapt As FirmwareVariable
    Dim loOpCapt As FirmwareVariable
    Dim loFwCmdExitPreOperation As FwCmdExitPreOperation
    Dim loFwCmdGotoOperation As FwCmdGotoOperation
        
    'DOPP-PATTERN
    '''-# Preconditioning measurement pattern at Vdd-5% using Diffmeter at AMUX to measure offsets for compensation
    lsPatNameExt = "Pre"
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' -# Program the DUT to the operation-mode, monitor the TX TOP states and configure the DPLL operating frequency
    Call DutPreopToOperation( _
        loFwCmdExitPreOperation, loFwCmdGotoOperation, True, , , True, loPreOpCapt, loLowPowCapt, loOpCapt)
    ''' -# Disable the DFX gate protection
    Call Ctrx.DisableDfxGateProtection
    ''' -# Monitor the chip temperature before starting measurement
    Call DtsMonitoringNoPattern(moFwCmdGetTempPre)
    ''' -# Disable the RX-subsystem
    Call RxRxChansDisabling(WAIT_100US)
    ''' -# Disable all Txrf phase shifters
    Call TxrfPsDisable
    ''' -# Disable all Txrf power amplifiers
    For liTxChannel = TXCHANNEL1 To TXCHANNEL8
        Call TxDisablePAChannels(liTxChannel, pbPaCfEnVld:=True)
    Next liTxChannel

    ''' -# Global Amux configured to connect PAD 1 and PAD 2 Differential to keep the PLD loads symmetric
    Call TxConfigGlobalAmux(OptAmuxPADConfType_Pad1En, OptAmuxConfType_Differential, _
                            OptAmuxPADConfType_Pad2En, OptAmuxConfType_Differential)

    ''' Step 7 Tx Supply voltage measurements - MIN supply levels
    Call TxgeMeasureSupplyLevels

    ''' Now step up the voltage supplies to maximum.
    Dopp.PatternPause
    Call TxSetVddSupply(VDDMAX_PERCENT)
    TheHdw.Wait (3 * ms)

    ''' Step 8 Tx Supply voltage measurements - MAX supply levels
    Call TxgeMeasureSupplyLevels
    ''' Pattern stop for Preconditioning results readback and Supply compensation calculations
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN

    msdPreOpDtct = loPreOpCapt.Value
    msdLowPowDtct = loLowPowCapt.Value
    msdOpDtct = loOpCapt.Value
    msdExitPreOpErr = loFwCmdExitPreOperation.Errorcode
    msdGotoOpErr = loFwCmdGotoOperation.Errorcode

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Routine to measure Vdd1V8 and Vdd1V0 supplies for no TxChannels enabled then for each TxChannel sequentially
'''@details
''' This subroutine must be called from within a Dopp block. Supply voltages are routed out to the AMUX pins.
''' Note that the Diffmeter instrument strobe is changed from AMUX2 to AMUX1 when measuring TxChannel 5 and higher.
'''
'''@log{2025-09-12, Channon Andrew, Initial Version}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeMeasureSupplyLevels()
    Dim liIdx As Long
    Dim lsMeasPin As String
    Dim lbTxCh1toTxCh4 As Boolean
    
    ''' Tx Supply voltage measurements
    ''' - Select no Tx channels, measure the VDD1V8A and VDD1V0TX via Tx1rf Amux without the PA enabled.
    Call TxTxrfSingleAmuxSel(TXCHANNEL1, pdVoltage:=VOLTAGE_1V8)
    Call Dopp.Settle(0.001)
    Call Dopp.DCDiffMeter("AMUX2P_DIFF").Strobe(mdPaDiffMeterWait)
    Call TxTxrfSingleAmuxSel(TXCHANNEL1, pdVoltage:=VOLTAGE_1V0)
    Call Dopp.Settle(0.002)
    Call Dopp.DCDiffMeter("AMUX2P_DIFF").Strobe(mdPaDiffMeterWait)
    Call TxTxrfSingleAmuxSel(TXCHANNEL1, pbReset:=True)
    
    ''' - Initialize loop settings then measure each Tx channel sequentially
    lsMeasPin = "AMUX2P_DIFF"
    lbTxCh1toTxCh4 = True
    
    For liIdx = TXCHANNEL1 To TXCHANNEL8
        If liIdx = TXCHANNEL5 Then
            lsMeasPin = "AMUX1P_DIFF"
            lbTxCh1toTxCh4 = False
        End If
        '''-# Enable the phase shifter of the current TXRF block and strobe the DC supply current measurement
        Call TxrfPsSingleEnable(liIdx)
        '''-# Enable the PA of the current TXRF block
        Call TxEnablePAChannels(liIdx, pbPaCfEnVld:=True)
        '''-# Configure the PA IDAC with setting 255
        Call TxrfRampPaDac(TxGetTxChanPair(liIdx), True, lbTxCh1toTxCh4, Not lbTxCh1toTxCh4)
        '''-# Measure the VDD1V8 and VDD1V0 supply voltages at the AMUX
        Call TxgeMeasureAmuxSupplyVoltages(liIdx, lsMeasPin, lbTxCh1toTxCh4)
    Next liIdx

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Run the Dopp pattern to measure the Vdd supplies.
'''@details
''' Measures the Vdd voltage for VDD0V9RF, VDD0V9PA and VDD1V8PA DCVIs using the AMUX. For the VddPA values,
''' one Tx channel is enabled, default TXCHANNEL1.
''' In total, 11 voltage measurements are made, all using AMUX1P_DIFF.
''' The function exits with the Dopp pattern paused
'''@param[in] psOfflineValues Path for filename to load offline measured values
'''@return[out] the measured Vdd supply voltages with 1 active Tx channel load conditions
'''
'''@log{2025-02-19, Channon Andrew, initial implementation @JIRA{CTRXTE-5553} }
'''@log{2025-04-11, Channon Andrew, revised readback for Dopp Read for STO @JIRA{CTRXTE-5553} }
'''@log{2025-09-11, Channon Andrew, revised for 8188\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsReadVddPre()
    
    Dim liIdx As Long
    Dim liVddLoop As Long
    Dim lvSite As Variant
    Dim lsPin As String
    Dim lsPin1 As String
    Dim lsPin2 As String
    Dim lpldVddMeas As New PinListData
    Dim lpldVddCorr As New PinListData
    Dim lpldVdd1V0pin1 As New PinListData
    Dim lpldVdd1V0pin2 As New PinListData
    Dim lpldVdd1V8pin1 As New PinListData
    Dim lpldVdd1V8pin2 As New PinListData
    
    ''' Initialize all variables for measurement and subsequent calculation
    lsPin1 = "AMUX1P_DIFF"
    lsPin2 = "AMUX2P_DIFF"
    Call mwVddCorrMin.CreateConstant(0, 18, DspDouble)
    Call mwVddCorrMax.CreateConstant(0, 18, DspDouble)
    
    For liVddLoop = 0 To 1    ' 0 = VddMin, 1 = VddMax
        ' Initialize variables for calculations
        If liVddLoop = 0 Then
            lpldVdd1V0pin1 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V0 * VDDMIN_PERCENT, lsPin1)
            lpldVdd1V0pin2 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V0 * VDDMIN_PERCENT, lsPin2)
            lpldVdd1V8pin1 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V8 * VDDMIN_PERCENT, lsPin1)
            lpldVdd1V8pin2 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V8 * VDDMIN_PERCENT, lsPin2)
        Else
            Set lpldVdd1V0pin1 = New PinListData
            Set lpldVdd1V0pin2 = New PinListData
            Set lpldVdd1V8pin1 = New PinListData
            Set lpldVdd1V8pin2 = New PinListData
            lpldVdd1V0pin1 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V0 * VDDMAX_PERCENT, lsPin1)
            lpldVdd1V0pin2 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V0 * VDDMAX_PERCENT, lsPin2)
            lpldVdd1V8pin1 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V8 * VDDMAX_PERCENT, lsPin1)
            lpldVdd1V8pin2 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V8 * VDDMAX_PERCENT, lsPin2)
        End If
        ' Loop for readback of uncompensated VDD measurements with calculation of compensation values
        ' - initialize pin name for no enabled channels (which uses TxCh1) and Tx channels 1-4
        lsPin = "AMUX2P_DIFF"
        For liIdx = 0 To 17
            If liIdx = 10 Then
                lsPin = "AMUX1P_DIFF"   ' for Tx channels 5-8
                Set lpldVddMeas = New PinListData
                Set lpldVddCorr = New PinListData
            End If
            lpldVddMeas = DoppRead(lsPin, Diffmeter)
            ' Calculate the correction factor based on the value measured
            ' then check against extreme voltage compensation levels in case something goes wrong
            If (liIdx Mod 2) = 0 Then
                ' even values of loop index for 1V8 measured values
                If liIdx < 10 Then
                    lpldVddCorr = lpldVdd1V8pin2.Math.Add(lpldVdd1V8pin2).Subtract(lpldVddMeas) 'TxCh 1-4
                Else
                    lpldVddCorr = lpldVdd1V8pin1.Math.Add(lpldVdd1V8pin1).Subtract(lpldVddMeas) 'TxCh 5-8
                End If
                lpldVddCorr = IfxMath.Calc.PinListData.Clip(lpldVddCorr, VDD1V8_THR_MIN, VDD1V8_THR_MAX)
            Else
                ' odd values of loop index for 1V0 measured values
                If liIdx < 10 Then
                    lpldVddCorr = lpldVdd1V0pin2.Math.Add(lpldVdd1V0pin2).Subtract(lpldVddMeas) 'TxCh 1-4
                Else
                    lpldVddCorr = lpldVdd1V0pin1.Math.Add(lpldVdd1V0pin1).Subtract(lpldVddMeas) 'TxCh 5-8
                End If
                lpldVddCorr = IfxMath.Calc.PinListData.Clip(lpldVddCorr, VDD1V0_THR_MIN, VDD1V0_THR_MAX)
            End If
            '''-# Convert the PinlistData values to DspWave, the format used to apply the levels
            If liVddLoop = 0 Then   ' store VddMin compensation values
                For Each lvSite In TheExec.Sites
                    mwVddCorrMin(lvSite).Element(liIdx) = lpldVddCorr.Pins(lsPin)(lvSite)
                Next lvSite
            Else    ' store VddMax compensation values
                For Each lvSite In TheExec.Sites
                    mwVddCorrMax(lvSite).Element(liIdx) = lpldVddCorr.Pins(lsPin)(lvSite)
                Next lvSite
            End If
        Next liIdx
    Next liVddLoop
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function configures the DCVI instruments for the main measurement pattern
'''@details
''' Supply voltages are configured for the Max corner measurements. Diffmeter setup is not changed - it keeps
''' the same settings as used in the preconditioning pattern
'''
'''@log{2025-04-11, Channon Andrew, Initial Version}
'''@log{2025-09-12, Channon Andrew, Adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsInstrSetup()
    Dim lsDcviPins As String

    '''-# Setup voltage supply DCVIs to measure current
    If (msHwSetTpConfig = "FE") Then
        lsDcviPins = "VDD1V8TX_DCVI, VDD1V8DPLL_DCVI, VDD1V0TX_MDCVI"
    Else
        lsDcviPins = "VDD1V8_DCVI, VDD1V0TX_MDCVI"
    End If
    TheHdw.Dcvi.Pins(lsDcviPins).Meter.Mode = tlDCVIMeterCurrent
    TheHdw.Dcvi.Pins(lsDcviPins).Meter.HardwareAverage = HW_AVERAGE_64
    TheHdw.Dcvi.Pins(lsDcviPins).Capture.SampleRate = 1 * MHz
    Call IfxHdw.Dcvi.BleederResistorOff(lsDcviPins)
    
    '''-# Update to the appropriate wait time for the DCVI pattern strobes
    mdDcviWait = IfxHdw.Dcvi.GetStrobeWait("VDD1V0TX_MDCVI")
    
    '''-# Set compensated supply voltages to Max corner +5%
    Call TxgeApplyVddCal(VDDMAX_PERCENT)
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Apply supply voltage compensation for Txge
'''@details
''' For Txge, the measured Vdd Calibration values are applied. A single DSPWave array contains the following values:
''' Selected TxChannel          0   1   2   3   4   5   6   7   8
''' Array Index for Vdd1V8A     0   2   4   6   8   10  12  14  16
''' Array Index for Vdd1V0TX    1   3   5   7   9   11  13  15  17
'''
'''@param[in] pdSelectMinMax Select the wanted setup: either for VDDMIN or VDDMAX
'''@param[in] piTxCh TX channel selected (optional)
'''@param[in] pdSettlingTime Settling time after changing DCVI voltage setting (optional)
'''
'''@log{2024-12-04, Channon Andrew, Initial Version\, @JIRA{CTRXTE-5360}}
'''@log{2025-09-12, Channon Andrew, Adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeApplyVddCal( _
    ByVal pdSelectMinMax As Double, _
    Optional ByVal piTxCh As Long = INITIALIZE0, _
    Optional ByVal pdSettlingTime As Double = -1)
    
    Dim ls1V8DcviPin As String
    
    If (msHwSetTpConfig = "FE") Then
        ls1V8DcviPin = "VDD1V8TX_DCVI"
    Else
        ls1V8DcviPin = "VDD1V8_DCVI"
    End If
    
    ''' Apply measured supply voltage compensation per Tx Channel (Also supports 0 Tx Channel)
    If pdSelectMinMax = VDDMIN_PERCENT Then
        For Each Site In TheExec.Sites
            Call IfxHdw.Dcvi.SetVoltage(ls1V8DcviPin, mwVddCorrMin(Site).Element(piTxCh * 2))
            Call IfxHdw.Dcvi.SetVoltage("VDD1V0TX_MDCVI", mwVddCorrMin(Site).Element((piTxCh * 2) + 1))
        Next Site
    ElseIf pdSelectMinMax = VDDMAX_PERCENT Then
        For Each Site In TheExec.Sites
            Call IfxHdw.Dcvi.SetVoltage(ls1V8DcviPin, mwVddCorrMax(Site).Element(piTxCh * 2))
            Call IfxHdw.Dcvi.SetVoltage("VDD1V0TX_MDCVI", mwVddCorrMax(Site).Element((piTxCh * 2) + 1))
        Next Site
    Else
        ' Wrong parameter value was entered. Exit with no action.
        Exit Sub
    End If
        
    If pdSettlingTime > 0 Then
        Call IfxHdw.Wait(pdSettlingTime)
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure the ICC currents from the Tx Supplies and measure the compensated supply voltages
'''@details
''' Run the pattern to measure the supply currents for VDD1V8TX and VDD1V0TX at minimum and maximum levels. Here we
''' also measure the compensated supply voltages using the Diffmeter via the AMUX.
''' Afterwards, read back the measured values and store for later processing and datalogging
''' Disconnect the DiffMeters after reading back the measured values to clear the setup for the next step.
'''
'''@log{2025-09-12, Channon Andrew, Initial Version}
'''@log{2025-09-12, Channon Andrew, Adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsMeasCorners()
    
    Dim lsPatNameExt As String
    Dim liIdx As Long
    Dim lsDiffMeterPin As String
    
    'DOPP-PATTERN
    lsPatNameExt = "Corners" & msPatHw
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' -# Enable the RX-subsystem
    Call RxRxChansEnabling(WAIT_100US)
    ''' Step 11. ICC offset measurements from the TX supplies with all Tx channels disabled and RX on at VDDMAX
    Call Dopp.Settle(WAIT_2MS)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' Step 12. Sequential ICC measurements from the TX supplies with only 1 TX channel enabled and RX on at VDDMAX
    Call TxgeCurConsMeasureTxPaCorners(VDDMAX_PERCENT)
    ''' Disable the RX-subsystem
    Call RxRxChansDisabling(WAIT_100US)
    ''' Pause pattern to set compensated supply levels to minimum with all TX channels disabled
    Call Ctrx.DoppPausePattern
    Call TxgeApplyVddCal(VDDMIN_PERCENT)
    ''' Step 13. ICC offset measurements from the TX supplies with all Tx channels disabled and RX on at VDDMIN
    Call Dopp.Settle(WAIT_2MS)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' Step 14. Sequential ICC measurements from the TX supplies with only 1 TX channel enabled and RX on at VDDMIN
    Call TxgeCurConsMeasureTxPaCorners(VDDMIN_PERCENT)
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN

    '''-# Read measured voltages and currents for Max corner settings
    lsDiffMeterPin = "AMUX2P_DIFF"
    mpldaIddMaxCorner(0) = DoppRead(msPinsCurrCons, Dcvi)       ' Baseline Offset current
    For liIdx = 0 To 7
        If liIdx = 4 Then
            lsDiffMeterPin = "AMUX1P_DIFF"
        End If
        mpldaIddMaxCorner(liIdx + 1) = DoppRead(msPinsCurrCons, Dcvi)       ' PS enabled
        mpldaIddMaxCorner(liIdx + 9) = DoppRead(msPinsCurrCons, Dcvi)       ' PA enabled
        mpldaVddMaxCorner(liIdx) = DoppRead(lsDiffMeterPin, Diffmeter)      ' Vdd1V8
        mpldaVddMaxCorner(liIdx + 8) = DoppRead(lsDiffMeterPin, Diffmeter)  ' Vdd1V0
    Next liIdx
    '''-# Read measured voltages and currents for Min corner settings
    lsDiffMeterPin = "AMUX2P_DIFF"
    mpldaIddMinCorner(0) = DoppRead(msPinsCurrCons, Dcvi)       ' Baseline Offset current
    For liIdx = 0 To 7
        If liIdx = 4 Then
            lsDiffMeterPin = "AMUX1P_DIFF"
        End If
        mpldaIddMinCorner(liIdx + 1) = DoppRead(msPinsCurrCons, Dcvi)       ' PS enabled
        mpldaIddMinCorner(liIdx + 9) = DoppRead(msPinsCurrCons, Dcvi)       ' PA enabled
        mpldaVddMinCorner(liIdx) = DoppRead(lsDiffMeterPin, Diffmeter)      ' Vdd1V8
        mpldaVddMinCorner(liIdx + 8) = DoppRead(lsDiffMeterPin, Diffmeter)  ' Vdd1V0
    Next liIdx

    ''' Disconnect DIFFMETERs
    Call TxAmuxDiffmeterSetup(DCVI_DIFF_DISC)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Monitor the DPLL 1V2 Bias voltage
'''
'''@log{2025-09-17, Channon Andrew, Initial version\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsDpllBias1V2()
    
    Dim lsPatNameExt As String
    Dim liTxIndex As Long
    
    ''' Measure bias voltage 1V2 for DPLL
    Set moRegTxDpllDftDfxgateRefCf1 = IfxRegMap.NewRegTxDpllDftDfxgateRefCf1
    Set moRegTxDpllSupplyDfxgateRefCf1 = IfxRegMap.NewRegTxDpllSupplyDfxgateRefCf1
    Set moRegTxDpllDcoCtrlDfxgateRefCf1 = IfxRegMap.NewRegTxDpllDcoCtrlDfxgateRefCf1
    
    ''' - Setup compensated supply voltages for Min (095) with no Tx channels enabled
    Call TxgeApplyVddCal(VDDMIN_PERCENT)
    ''' DCDiffMeter configuration for differential measurements on the AMUXes
    ''' - High Accuracy mode, Voltage range 3.5V, HW_AVERAGE = 64, bleeder resistors OFF
    Call IfxHdw.DCDiffMeter.Setup( _
        "AMUX2N_DIFF", "AMUX2P_DIFF", tlDCDiffMeterModeHighAccuracy, DiffMeterRange3p5V, HW_AVERAGE_64, True)
    Call IfxHdw.Dcvi.BleederResistorOff("AMUX2_P_DCVI,AMUX2_N_DCVI")
        
    ''' Get the wait time for diffmeter measurements
    mdPaDiffMeterWait = IfxHdw.DCDiffMeter.GetStrobeWait("AMUX2N_DIFF") * 2
    
    'DOPP-PATTERN
    lsPatNameExt = "DpllBias1V2" '& msPatHw
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' Connect the DPLL_AMUX output to the TG_AMUX_SADC
    With IfxRegMap.NewRegTxTxamuxTxAmuxSadcChCf
        .bfSel = OptTxTxraigroupTxamuxTxAmuxSadcChConf_Sel_0x0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    moRegTxDpllDftDfxgateRefCf1.bfDftEn = 1
    moRegTxDpllDftDfxgateRefCf1.bfDftBufEn = 1
    moRegTxDpllDftDfxgateRefCf1.bfDftSel = 1
    moRegTxDpllDftDfxgateRefCf1.bfDftSelB = 0
    moRegTxDpllDftDfxgateRefCf1.bfMmdLdoDftEn = 1
    moRegTxDpllDftDfxgateRefCf1.bfDtcLdoDftEn = 1
    moRegTxDpllDftDfxgateRefCf1.bfTdcLdoDftEn = 1
    Call Ctrx.WriteFixedRegister(moRegTxDpllDftDfxgateRefCf1)
    moRegTxDpllDcoCtrlDfxgateRefCf1.bfDcoDftEn = 1
    Call Ctrx.WriteFixedRegister(moRegTxDpllDcoCtrlDfxgateRefCf1)
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxEn = 1
    moRegTxDpllSupplyDfxgateRefCf1.bfLsLdoDftEn = 1
    'Set the Amux control value to route out DPLL 1V2 Bias (channel A) and VSSA (channel B)
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxCtrl = 6
    Call Ctrx.WriteFixedRegister(moRegTxDpllSupplyDfxgateRefCf1)
    Call Dopp.Settle(WAIT_2MS)
    Dopp.DCDiffMeter("AMUX2N_DIFF").Strobe (mdPaDiffMeterWait)
    ''' Route the TG_AMUX_SADC from the TX_BIAS_AMUX towards the AMUX1 so that the AMUX2 works with the TXRF AMUXes
    With IfxRegMap.NewRegTxTxamuxTxAmuxSadcChCf
        .bfSel = OptTxTxraigroupTxamuxTxAmuxSadcChConf_Sel_0x2
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Ctrx.WriteFixedRegisterWithComment(IfxRegMap.NewRegTxDpllDftDfxgateRefCf1, Logger)
    Call Ctrx.WriteFixedRegisterWithComment(IfxRegMap.NewRegTxDpllDcoCtrlDfxgateRefCf1, Logger)
    Call Ctrx.WriteFixedRegisterWithComment(IfxRegMap.NewRegTxDpllSupplyDfxgateRefCf1, Logger)
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN
    
    ''' Read the bias voltage 1V2 for DPLL measurement from the DCDiffMeter and store them for datalogging
    If TheExec.TesterMode = testModeOnline Then
        mpldDpllSense1V2 = DoppRead("AMUX2N_DIFF", Diffmeter)
        mpldDpllSense1V2 = mpldDpllSense1V2.Math.Abs
    End If

    ''' Reset the DCDiffMeter
    Call IfxHdw.DCDiffMeter.Reset("AMUX2N_DIFF", True, True)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure the DPLL Bias and IBG currents
'''@details
''' Must be called from inside a Dopp pattern block
'''
'''@log{2025-09-17, Channon Andrew, Initial version\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsDpllBiasCurrents()
    
    ''' Measure bias voltage 1V2 for DPLL
    Set moRegTxDpllDftDfxgateRefCf1 = IfxRegMap.NewRegTxDpllDftDfxgateRefCf1
    Set moRegTxDpllSupplyDfxgateRefCf1 = IfxRegMap.NewRegTxDpllSupplyDfxgateRefCf1
    Set moRegTxDpllDcoCtrlDfxgateRefCf1 = IfxRegMap.NewRegTxDpllDcoCtrlDfxgateRefCf1
    
    ''' Connect the DPLL_AMUX output to the TG_AMUX_SADC
    With IfxRegMap.NewRegTxTxamuxAmuxCh2Cf
        .bfAmuxAx2pCtrl = OptTxTxraigroupTxamuxAmuxCh2Conf_AmuxAx2pCtrl_0x4
        .bfAmuxAx2nCtrl = OptTxTxraigroupTxamuxAmuxCh2Conf_AmuxAx2nCtrl_0x4
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    With IfxRegMap.NewRegTxTxamuxTxAmuxSadcChCf
        .bfSel = OptTxTxraigroupTxamuxTxAmuxSadcChConf_Sel_0x0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    moRegTxDpllDftDfxgateRefCf1.bfDftEn = 1
    moRegTxDpllDftDfxgateRefCf1.bfDftBufEn = 0
    moRegTxDpllDftDfxgateRefCf1.bfDftSel = 0
    moRegTxDpllDftDfxgateRefCf1.bfDftSelB = 1
    Call Ctrx.WriteFixedRegister(moRegTxDpllDftDfxgateRefCf1)
    
    ''' 50uA DCO bias current (left current mirror)
    moRegTxDpllDcoCtrlDfxgateRefCf1.bfDcoTstCurrent1En = 1
    Call Ctrx.WriteFixedRegister(moRegTxDpllDcoCtrlDfxgateRefCf1)
    Call Dopp.Settle(WAIT_5MS)
    Call Dopp.Dcvi.Pins("AMUX2_P_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' 50uA DCO bias current (right current mirror)
    moRegTxDpllDcoCtrlDfxgateRefCf1.bfDcoTstCurrent1En = 0
    moRegTxDpllDcoCtrlDfxgateRefCf1.bfDcoTstCurrent2En = 1
    Call Ctrx.WriteFixedRegister(moRegTxDpllDcoCtrlDfxgateRefCf1)
    Call Dopp.Settle(WAIT_2MS)
    Call Dopp.Dcvi.Pins("AMUX2_P_DCVI").Meter.Strobe(mdDcviWait)
    moRegTxDpllDcoCtrlDfxgateRefCf1.bfDcoTstCurrent2En = 0
    Call Ctrx.WriteFixedRegister(moRegTxDpllDcoCtrlDfxgateRefCf1)
    
    ''' IBG 10uA DPLL current monitoring
    moRegTxDpllDftDfxgateRefCf1.bfDftSelB = 2
    Call Ctrx.WriteFixedRegister(moRegTxDpllDftDfxgateRefCf1)
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxEn = 1
    moRegTxDpllSupplyDfxgateRefCf1.bfLsLdoDftEn = 1
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxCtrl = 1
    Call Ctrx.WriteFixedRegister(moRegTxDpllSupplyDfxgateRefCf1)
    Call Dopp.Settle(WAIT_100US)
    Call Dopp.Dcvi.Pins("AMUX2_P_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' IBG 50uA DPLL current monitoring
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxCtrl = 2
    Call Ctrx.WriteFixedRegister(moRegTxDpllSupplyDfxgateRefCf1)
    Call Dopp.Settle(WAIT_100US)
    Call Dopp.Dcvi.Pins("AMUX2_P_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' IBG 10uA DPLL PTAT current monitoring
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxCtrl = 3
    Call Ctrx.WriteFixedRegister(moRegTxDpllSupplyDfxgateRefCf1)
    Call Dopp.Settle(WAIT_100US)
    Call Dopp.Dcvi.Pins("AMUX2_P_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' IBG 20uA DPLL PTAT current monitoring
    moRegTxDpllSupplyDfxgateRefCf1.bfBiasAmuxCtrl = 4
    Call Ctrx.WriteFixedRegister(moRegTxDpllSupplyDfxgateRefCf1)
    Call Dopp.Settle(WAIT_100US)
    Call Dopp.Dcvi.Pins("AMUX2_P_DCVI").Meter.Strobe(mdDcviWait)
    
    ' reset
    ''' Route the TG_AMUX_SADC from the TX_BIAS_AMUX towards the AMUX1 so that the AMUX2 works with the TXRF AMUXes
    With IfxRegMap.NewRegTxTxamuxTxAmuxSadcChCf
        .bfSel = OptTxTxraigroupTxamuxTxAmuxSadcChConf_Sel_0x2
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Ctrx.WriteFixedRegisterWithComment(IfxRegMap.NewRegTxDpllDftDfxgateRefCf1, Logger)
    Call Ctrx.WriteFixedRegisterWithComment(IfxRegMap.NewRegTxDpllDcoCtrlDfxgateRefCf1, Logger)
    Call Ctrx.WriteFixedRegisterWithComment(IfxRegMap.NewRegTxDpllSupplyDfxgateRefCf1, Logger)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure the TX Bias and TXLO IBG currents
'''@details
''' Must be called from inside a Dopp pattern block
'''
'''@log{2025-09-17, Channon Andrew, Initial version\, @JIRA{RSIPPTE-109}}
'''@log{2025-10-28, Channon Andrew, fix STO error with pin names\, @JIRA{RSIPPTE-410}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsTxBiasCurrents()
    
    ''' trimmed IBG 50uA
    With IfxRegMap.NewRegTxTxbiasSenseCf
        .bfAmuxCtrl = OptTxTxraigroupTxbiasSenseConf_AmuxCtrl_0x5
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(250 * us)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' untrimmed IBG 50uA
    With IfxRegMap.NewRegTxTxbiasSenseCf
        .bfAmuxCtrl = OptTxTxraigroupTxbiasSenseConf_AmuxCtrl_0x6
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(100 * us)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' trimmed 50uA PTAT
    With IfxRegMap.NewRegTxTxbiasSenseCf
        .bfAmuxCtrl = OptTxTxraigroupTxbiasSenseConf_AmuxCtrl_0x7
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(100 * us)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' Reset Amux
    With IfxRegMap.NewRegTxTxbiasSenseCf
        .bfAmuxCtrl = OptTxTxraigroupTxbiasSenseConf_AmuxCtrl_0x0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    
    ''' LO DIST IBG 20uA
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x12
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(100 * us)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' Reset Amux
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure the TXLO Bias currents
'''@details
''' Must be called from inside a Dopp pattern block
'''
'''@log{2025-09-17, Channon Andrew, Initial version\, @JIRA{RSIPPTE-109}}
'''@log{2025-10-28, Channon Andrew, fix STO error with pin names\, @JIRA{RSIPPTE-410}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsTxloBiasCurrents()

    ''' LOOUT TC1 20uA current
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x23
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(2 * ms)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' LOIN TC2 20uA current
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x24
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(2 * ms)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' FQM2 TC3 20uA current
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x25
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(2 * ms)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' FQM3 TC4 20uA current
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x26
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(2 * ms)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ''' CASSPLIT-COMB TC5 20uA current
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x27
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Settle(2 * ms)
    Call Dopp.Dcvi.Pins("AMUX1_P_DCVI,AMUX1_N_DCVI").Meter.Strobe(mdDcviWait)
    
    ' reset
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief TX DC current measurements
'''
'''@log{2025-09-17, Andrew Channon, Initial Version }
'''@log{2025-09-17, Andrew Channon, Ported from 8191 and adapted for 8188 \, @JIRA{RSIPPTE-109}}
'''@log{2025-10-21, Andrew Channon, Adjust DCVI instrument settings to align to REFX \, @JIRA{RSIPPTE-401}}
'''@log{2025-10-28, Channon Andrew, fix STO error with pin names\, @JIRA{RSIPPTE-410}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsDcCurrMeas()
    
    Dim lsPatNameExt As String
    Dim lsPins As String
    Dim liIdx As Long
    Dim lsMeasPins As String
    
    lsPins = IfxMath.Util.String_.Join(",", "AMUX1_P_DCVI", "AMUX1_N_DCVI", "AMUX2_P_DCVI", "AMUX2_N_DCVI")
    
    '''-# Setup the DCVI meters on AMUX1P and AMUX2P for current measurement
    Call IfxHdw.Dcvi.Setup(lsPins, tlDCVIModeVoltage, FORCE_V_NMOS_VAL, 200 * uA, DcviRange200uA, _
        peMeterMode:=tlDCVIMeterCurrent, peBleederSetting:=tlDCVIBleederResistorOff)
    
    '''-# Get the wait time for DCVI setup measurement.
    mdDcviWait = IfxHdw.Dcvi.GetStrobeWait("AMUX1_P_DCVI")
    
    'DOPP-PATTERN
    lsPatNameExt = "DcCurrents" '& msPatHw
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' Step 19. Initialize loop settings for TXRF IBG current measurements then measure each Tx channel sequentially
    lsMeasPins = "AMUX2_P_DCVI,AMUX2_N_DCVI"
    For liIdx = TXCHANNEL1 To TXCHANNEL8
        If liIdx = TXCHANNEL5 Then
            lsMeasPins = "AMUX1_P_DCVI,AMUX1_N_DCVI"
        End If
        '''-# Measure the untrimmed IBG 20uA at the AMUX
        Call TxTxrfSingleAmuxSelCurrent(liIdx, pdCurrent:=CURRENT_IBG20UA)
        Call Dopp.Settle(100 * us)
        Call Dopp.Dcvi.Pins(lsMeasPins).Meter.Strobe(mdDcviWait)
        '''-# Measure the trimmed IBG 10uA at the AMUX
        Call TxTxrfSingleAmuxSelCurrent(liIdx, pdCurrent:=CURRENT_IBG10UA)
        Call Dopp.Settle(100 * us)
        Call Dopp.Dcvi.Pins(lsMeasPins).Meter.Strobe(mdDcviWait)
        '''-# Reset the AMUX of the current TXRF block
        Call TxTxrfSingleAmuxSelCurrent(liIdx, pbReset:=True)
    Next liIdx
    ''' Step 20. Tx Bias and TXLO current measurements
    Call TxgeCurConsTxBiasCurrents
    ''' Step 21. DPLL bias and IBG currents
    Call TxgeCurConsDpllBiasCurrents
    ''' Step 22. TXLO bias measurements
    Call TxgeCurConsTxloBiasCurrents
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN
    
    '''-# Read DC currents
    '''--# Read Txrf Ibg currents
    lsMeasPins = "AMUX2_P_DCVI,AMUX2_N_DCVI"
    For liIdx = 0 To 7
        If liIdx = 4 Then
            lsMeasPins = "AMUX1_P_DCVI,AMUX1_N_DCVI"
        End If
        mpldaTxrfIbgCurr(liIdx) = DoppRead(lsMeasPins, Dcvi)
        mpldaTxrfIbgCurr(liIdx + 8) = DoppRead(lsMeasPins, Dcvi)
    Next liIdx
    '''--# Read Tx Bias currents
    For liIdx = 0 To 3
        mpldaTxBiasCurr(liIdx) = DoppRead("AMUX1_P_DCVI,AMUX1_N_DCVI", Dcvi)
    Next liIdx
    '''--# Read Dpll currents
    For liIdx = 0 To 5
        mpldaDpllBiasCurr(liIdx) = DoppRead("AMUX2_P_DCVI", Dcvi)
    Next liIdx
    '''--# Read Txlo Bias currents
    For liIdx = 0 To 4
        mpldaTxloBiasCurr(liIdx) = DoppRead("AMUX1_P_DCVI,AMUX1_N_DCVI", Dcvi)
    Next liIdx
    
    '''-# Now reset DCVIs to release them for Diffmeter usage
    Call IfxHdw.Dcvi.Reset(lsPins)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure the Tx supply current consumption with the DIV8 enabled
'''@details
''' The pattern using the common DIV8 enabling sequence defined in Util_Tx. As this is the last test in the TX current
''' consumption measurements, the pattern continues with the DUT deconditioning steps. It monitors the DTS for the
''' chip temperature after measurements then re-enables the DfxGateProtection before resetting the device.
''' The measured supply currents with the DIV8 enabled are read back here.
'''
'''@log{2025-09-19, Channon Andrew, Initial Version}
'''@log{2025-10-28, Channon Andrew, Remove Div8 tests\, @JIRA{RSIPPTE-411}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsReset()
    Dim lsPatNameExt As String
       
    'DOPP-PATTERN
    lsPatNameExt = "Reset"
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' -# Monitor the DTS after the measurements and before resetting the DUT
    Call DtsMonitoringNoPattern(moFwCmdGetTempPost)
    ''' -# Step 26. Enable back the DFX protection
    Call Ctrx.EnableDfxGateProtection
    ''' -# Step 27. Soft-Reset and pattern stop
    Call Ctrx.SoftReset(peUpdateMethod:=ForceUpdate, pbAlarmToggleRequired:=True)
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure corner supply currents and supply voltages for each TXRF channel sequentially
'''@details
''' Must be executed within a Dopp block. Each TXRF channel is enabled in turn. Supply current is measured after
''' enabling the Phase Shifter, The supply current is enabled again after enabling the PA and setting the IDAC to 255
''' At this point, both the VDD1V8 and VDD1V0 supply voltages are measured using the AMUX when the TXRF channel is
''' fully loaded. Each TXRF channel is then fully reset before moving to the next TXRF channel for measurement.
'''
'''@param[in] pdSelectMinMax Select the wanted setup: either for VDDMIN or VDDMAX
'''
'''@log{2022-03-03, Vommi Suneel, Initial Version}
'''@log{2022-03-22, Vommi Suneel, Adopted Voltage Correction  @JIRA{CTRXTE-1001}}
'''@log{2022-04-25, Vommi Suneel, Adopted Voltage Correction  @JIRA{CTRXTE-1082}}
'''@log{2023-05-29, Vommi Suneel, Update B-step  Tp\, @JIRA{CTRXTE-1898}}
'''@log{2023-05-29, Vommi Suneel, Debug Update B-step  Tp\, @JIRA{CTRXTE-1899}}
'''@log{2023-07-12, Vommi Suneel, Test-0 Trail run fix B-step  Tp\, @JIRA{CTRXTE-2880}}
'''@log{2024-09-16, Vommi Suneel, Splitting CAS Stand  Tp\, @JIRA{CTRXTE-5016}}
'''@log{2024-11-12, Neul Roland, move SPI/JTAG support to library\, @JIRA{CTRXTE-5202}}
'''@log{2025-02-13, Channon Andrew, Optimize for STO TTR\, @JIRA{CTRXTE-5553}}
'''@log{2025-09-15, Channon Andrew, Refactored for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsMeasureTxPaCorners( _
    ByVal pdSelectMinMax As Double)

    Dim liIdx As Long
    Dim lsMeasPin As String
    Dim lbTxCh1toTxCh4 As Boolean
    
    ''' Initialize loop settings then measure each Tx channel sequentially
    lsMeasPin = "AMUX2P_DIFF"
    lbTxCh1toTxCh4 = True
    
    For liIdx = TXCHANNEL1 To TXCHANNEL8
        If liIdx = TXCHANNEL5 Then
            lsMeasPin = "AMUX1P_DIFF"
            lbTxCh1toTxCh4 = False
        End If
        '''-# Pause pattern to set compensated supply levels for the selected TX channel
        Call Ctrx.DoppPausePattern
        Call TxgeApplyVddCal(pdSelectMinMax, liIdx)
        '''-# Enable the phase shifter of the current TXRF block and strobe the DC supply current measurement
        Call TxrfPsSingleEnable(liIdx)
        Call Dopp.Settle(WAIT_1MS)
        Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
        '''-# Enable the PA of the current TXRF block
        Call TxEnablePAChannels(liIdx, pbPaCfEnVld:=True)
        '''-# Configure the PA IDAC with setting 255 then strobe the DC supply current measurement
        Call TxrfRampPaDac(TxGetTxChanPair(liIdx), True, lbTxCh1toTxCh4, Not lbTxCh1toTxCh4)
        Call Dopp.Settle(WAIT_1MS)
        Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
        '''-# Measure the VDD1V8 and VDD1V0 supply voltages at the AMUX
        Call TxgeMeasureAmuxSupplyVoltages(liIdx, lsMeasPin, lbTxCh1toTxCh4)
    Next liIdx

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This Function calculates the delta current
'''@details
'''Calculate the delta current of system
'''
'''@param[in] ppldIddTxSys Tx chain subsystem delta current
'''@param[in] ppldIddSys  Measured System current
'''@return[out] lpldDelta current as PinListData
'''
'''@log{2025-09-16, Channon Andrew, Ported from 8191\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function TxgeTxSysDeltaCurr( _
    ByRef ppldIddTxSys As PinListData, _
    ByRef ppldIddSys As PinListData _
    ) As PinListData
    
    Dim lpldDelta As New PinListData
    Dim lsDomain As String
    Dim liDomain As Integer
    
    '''- Calculate TX chain Delta Current
    For liDomain = 0 To ppldIddSys.Pins.Count - 1
        lsDomain = ppldIddTxSys.Pins(liDomain).Name
        lpldDelta.AddPin(lsDomain).Value = ppldIddTxSys.Pins(lsDomain).Subtract(ppldIddSys.Pins(lsDomain))
    Next liDomain

    Set TxgeTxSysDeltaCurr = lpldDelta
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This Function calculates the sum current
'''@details
'''Calculate the sum current of 2 pins
'''
'''@param[in] ppldIddTxSys Tx chain subsystem delta current
'''@return[out] lpldSum current as PinListData
'''
'''@log{2025-09-19, Channon Andrew, Initial version\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function TxgeTxSysSumCurr( _
    ByRef ppldIddTxSys As PinListData _
    ) As PinListData
    
    Dim lpldSum As New PinListData
    Dim lsDomain As String
    
    '''- Calculate TX Sum Current
    lsDomain = ppldIddTxSys.Pins(0).Name
    If ppldIddTxSys.Pins.Count = 2 Then
        lpldSum.AddPin(lsDomain).Value = ppldIddTxSys.Pins(0).Add(ppldIddTxSys.Pins(1))
    Else
        lpldSum.AddPin(lsDomain).Value = -0.999
    End If

    Set TxgeTxSysSumCurr = lpldSum
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Calculate the TX current consumption values from all measured supply values
'''
'''@log{2025-09-16, Channon Andrew, Initial Version adapted from 8191\, @JIRA{RSIPPTE-109}}
'''@log{2025-10-28, Channon Andrew, Remove Div8 tests\, @JIRA{RSIPPTE-411}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsCalc()
    Dim liIdx As Long
    Dim lsdIccDeltaPa As New SiteDouble
    
    ''' Idd Max Corner currents
    mpldaCurrResults(0) = mpldaIddMaxCorner(0)  ' Baseline Offset
    For liIdx = 1 To 8
        mpldaCurrResults(liIdx) = TxgeTxSysDeltaCurr(mpldaIddMaxCorner(liIdx), mpldaIddMaxCorner(0))             ' PS
        mpldaCurrResults(liIdx + 8) = TxgeTxSysDeltaCurr(mpldaIddMaxCorner(liIdx + 8), mpldaIddMaxCorner(liIdx)) ' PA
        lsdIccDeltaPa = mpldaCurrResults(liIdx + 8).Pins("VDD1V0TX_MDCVI")
        msdIccVdd1V0MaxCornerSum = msdIccVdd1V0MaxCornerSum.Add(lsdIccDeltaPa)
    Next liIdx
    ''' Idd Min Corner currents
    mpldaCurrResults(17) = mpldaIddMinCorner(0)  ' Baseline Offset
    For liIdx = 1 To 8
        mpldaCurrResults(liIdx + 17) = TxgeTxSysDeltaCurr(mpldaIddMinCorner(liIdx), mpldaIddMinCorner(0))         ' PS
        mpldaCurrResults(liIdx + 25) = TxgeTxSysDeltaCurr(mpldaIddMinCorner(liIdx + 8), mpldaIddMinCorner(liIdx)) ' PA
        lsdIccDeltaPa = mpldaCurrResults(liIdx + 25).Pins("VDD1V0TX_MDCVI")
        msdIccVdd1V0MinCornerSum = msdIccVdd1V0MinCornerSum.Add(lsdIccDeltaPa)
    Next liIdx
    ''' Txrf Ibg currents
    For liIdx = 0 To 15
        ' only pull the -P_DCVI values for single-ended datalog. Ibg current is the same on -P and -N DCVIs
        mpldaCurrResults(liIdx + 34) = mpldaTxrfIbgCurr(liIdx).Copy
    Next liIdx
    ''' Tx Bias and Txlo Ibg
    For liIdx = 0 To 3
        mpldaCurrResults(liIdx + 50) = TxgeTxSysSumCurr(mpldaTxBiasCurr(liIdx))
    Next liIdx
    ''' Txlo Bias currents
    For liIdx = 0 To 4
        mpldaCurrResults(liIdx + 54) = TxgeTxSysSumCurr(mpldaTxloBiasCurr(liIdx))
    Next liIdx
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This Function Datalog the sub system current
'''
'''@param[in] ppldDeltaI Subsystem current
'''@param[in] psPin1 Subsystem current
'''@param[in] psPin2 Subsystem current
'''@param[in] psPin3 Subsystem current
'''
'''@log{2022-04-21, Velummylum Mathiy, Initial Version}
'''@log{2023-05-02, Gallo Elisa, added optional parameters\, @JIRA{CTRXTE-2160}}
'''@log{2025-09-16, Channon Andrew, ported from 8191 and adapted for 8188\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeDcCurrDatalog( _
    ByRef ppldDeltaI As PinListData, _
    Optional psPin1 As String, _
    Optional psPin2 As String, _
    Optional psPin3 As String)

    If Len(psPin1) & Len(psPin2) & Len(psPin3) = 0 Then
        '''- Datalog of Tx SubSystem current
        Call IfxEnv.Datalog(ppldDeltaI)
    Else
        If Len(psPin1) <> 0 Then
            Call IfxEnv.Datalog(ppldDeltaI.Pins(psPin1))
        End If
        If Len(psPin2) <> 0 Then
            Call IfxEnv.Datalog(ppldDeltaI.Pins(psPin2))
        End If
        If Len(psPin3) <> 0 Then
            Call IfxEnv.Datalog(ppldDeltaI.Pins(psPin3))
        End If
    End If
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This Function Datalog the sub system current
'''
'''@param[in] ppldaResults() Subsystem current results array
'''@param[in] psPin1 Subsystem current
'''@param[in] psPin2 Subsystem current
'''@param[in] psPin3 Subsystem current
'''@param[in] piStart Array index to start datalog from
'''@param[in] piSize Number of items to datalog
'''
'''@log{2022-04-21, Velummylum Mathiy, Initial Version}
'''@log{2023-05-02, Gallo Elisa, added optional parameters\, @JIRA{CTRXTE-2160}}
'''@log{2025-09-16, Channon Andrew, ported from 8191 and adapted for 8188\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeDcCurrDatalogArray( _
    ByRef ppldaResults() As PinListData, _
    Optional psPin1 As String, _
    Optional psPin2 As String, _
    Optional psPin3 As String, _
    Optional piStart As Long = 0, _
    Optional piSize As Long = 1)

    Dim liIdx As Long
    Dim liEnd As Long
    liEnd = piStart + (piSize - 1)
    
    If Len(psPin1) & Len(psPin2) & Len(psPin3) = 0 Then
        For liIdx = piStart To liEnd
            '''- Datalog of Tx SubSystem current
            Call IfxEnv.Datalog(ppldaResults(liIdx))
        Next liIdx
    Else
        If Len(psPin1) <> 0 Then
            For liIdx = piStart To liEnd
                Call IfxEnv.Datalog(ppldaResults(liIdx).Pins(psPin1))
            Next liIdx
        End If
        If Len(psPin2) <> 0 Then
            For liIdx = piStart To liEnd
                Call IfxEnv.Datalog(ppldaResults(liIdx).Pins(psPin2))
            Next liIdx
        End If
        If Len(psPin3) <> 0 Then
            For liIdx = piStart To liEnd
                Call IfxEnv.Datalog(ppldaResults(liIdx).Pins(psPin3))
            Next liIdx
        End If
    End If
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Define the supply pins for datalogging current according to FE or BE selection
'''@param[out] psVdd1v8Pin1 Pin assignment for datalogging
'''@param[out] psVdd1v0Pin2  Pin assignment for datalogging
'''@param[out] psVdd1v8DllPin3  Pin assignment for datalogging
'''
'''@log{2025-09-16, Channon Andrew, Initial creation\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsDatalogPins( _
    ByRef psVdd1v8Pin1 As String, _
    ByRef psVdd1v0Pin2 As String, _
    ByRef psVdd1v8DpllPin3 As String)
    
    '''- Define the supply pins for datalogging current
    If (msHwSetTpConfig = "FE") Then
        psVdd1v8Pin1 = "VDD1V8TX_DCVI"
        psVdd1v8DpllPin3 = "VDD1V8DPLL_DCVI"
    Else
        psVdd1v8Pin1 = "VDD1V8_DCVI"
        psVdd1v8DpllPin3 = vbNullString
    End If
    psVdd1v0Pin2 = "VDD1V0TX_MDCVI"

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Datalog all results for the Txge Current Consumption tests
'''
'''@log{2025-09-15, Channon Andrew, Initial Version}
'''@log{2025-10-28, Channon Andrew, Remove Div8 tests\, @JIRA{RSIPPTE-411}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeCurConsDatalog()
    Dim liIdx As Long
    Dim lsVdd1v8Pin1 As String
    Dim lsVdd1v0Pin2 As String
    Dim lsVdd1v8DpllPin3 As String
    
    Call TxgeCurConsDatalogPins(lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3)
    ''' TX pre-conditioning status monitoring values
    Call IfxEnv.Datalog(msdPreOpDtct)
    Call IfxEnv.Datalog(msdLowPowDtct)
    Call IfxEnv.Datalog(msdOpDtct)
    Call IfxEnv.Datalog(msdExitPreOpErr)
    Call IfxEnv.Datalog(msdGotoOpErr)
    ''' Chip temperature readouts before and after the measurements
    Call IfxEnv.Datalog(moFwCmdGetTempPre.Temperature(DTS_TEMP1))
    Call IfxEnv.Datalog(moFwCmdGetTempPre.Temperature(DTS_TEMP2))
    Call IfxEnv.Datalog(moFwCmdGetTempPost.Temperature(DTS_TEMP1))
    Call IfxEnv.Datalog(moFwCmdGetTempPost.Temperature(DTS_TEMP2))
    ''' Vdd Max Corner supply voltages
    For liIdx = 0 To 15
        Call IfxEnv.Datalog(mpldaVddMaxCorner(liIdx))
    Next liIdx
    ''' Idd Max Corner currents
    Call TxgeDcCurrDatalog(mpldaCurrResults(0), lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3)
    ''' - PS currents
    Call TxgeDcCurrDatalogArray(mpldaCurrResults, lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3, piStart:=1, piSize:=8)
    ''' - PA currents
    Call TxgeDcCurrDatalogArray(mpldaCurrResults, lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3, piStart:=9, piSize:=8)
    ''' Sum PA currents Max Vdd
    Call IfxEnv.Datalog(msdIccVdd1V0MaxCornerSum)
    ''' Vdd Min Corner supply voltages
    For liIdx = 0 To 15
        Call IfxEnv.Datalog(mpldaVddMinCorner(liIdx))
    Next liIdx
    ''' Idd Min Corner currents
    Call TxgeDcCurrDatalog(mpldaCurrResults(17), lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3)
    ''' - PS currents
    Call TxgeDcCurrDatalogArray(mpldaCurrResults, lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3, piStart:=18, piSize:=8)
    ''' - PA currents
    Call TxgeDcCurrDatalogArray(mpldaCurrResults, lsVdd1v8Pin1, lsVdd1v0Pin2, lsVdd1v8DpllPin3, piStart:=26, piSize:=8)
    ''' Sum PA currents Min Vdd
    Call IfxEnv.Datalog(msdIccVdd1V0MinCornerSum)
    ''' Dpll 1V2
    Call IfxEnv.Datalog(mpldDpllSense1V2)
    ''' TXRF channels Ibg20uA and Ibg10uA
    For liIdx = 34 To 49
        Call IfxEnv.Datalog(mpldaCurrResults(liIdx).Pins(0))
    Next liIdx
    ''' Step 20. Tx Bias and TXLO current measurements
    Call TxgeDcCurrDatalogArray(mpldaCurrResults, piStart:=50, piSize:=4)
    ''' Step 21. DPLL bias and IBG currents
    Call TxgeDcCurrDatalogArray(mpldaDpllBiasCurr, piStart:=0, piSize:=6)
    ''' Step 22. TXLO bias measurements
    Call TxgeDcCurrDatalogArray(mpldaCurrResults, piStart:=54, piSize:=5)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function initializes the variables defined at module level
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsVariablesInit()
    Dim lsPins As String
    Dim lsVdd1v8Pin As String
    Dim liCount As Long
        
    If IfxStd.TesterHw.IsFrontendHw Then
        lsVdd1v8Pin = "VDD1V8TX_DCVI"
    Else
        lsVdd1v8Pin = "VDD1V8_DCVI"
    End If
    
    msHwSetTpConfig = IfxRadar.TpConfig("HWSETUP")
    
    lsPins = IfxMath.Util.String_.Join(",", lsVdd1v8Pin, "VDD1V0TX_MDCVI")
    If (msHwSetTpConfig = "FE") Then
        msPatHw = "FE"
    Else
        msPatHw = "BE"
    End If

    Set msdPreOpDtct = Nothing
    Set msdLowPowDtct = Nothing
    Set msdOpDtct = Nothing
    Set msdExitPreOpErr = Nothing
    Set msdGotoOpErr = Nothing
    Set moFwCmdGetTempPre = Nothing
    Set moFwCmdGetTempPost = Nothing
    
    Set mpldVldInFqm2Offset = Nothing
    Set mpldVldOutFqm2Offset = Nothing
    Set mpldVldInFqm3Offset = Nothing
    Set mpldVldOutFqm3Offset = Nothing
    Set mpldVldLoin1Offset = Nothing
    Set mpldVldLoin2Offset = Nothing
    Set mpldVldLoout1Offset = Nothing
    Set mpldVldLoout2Offset = Nothing
    
    Set mpldIccStandalone = Nothing
    Set mpldIccTxmonOff = Nothing
    Set mpldIccLoDistOff = Nothing
    Set mpldVldInFqm2Standalone = Nothing
    Set mpldVldOutFqm2Standalone = Nothing
    Set mpldVldInFqm3Standalone = Nothing
    Set mpldVldOutFqm3Standalone = Nothing
    Set mpldIccLo1StaOff = Nothing
    Set mpldIccLo2StaOn = Nothing
    Set mpldIccLo1Path = Nothing
    Set mpldIccLoin1Off = Nothing
    Set mpldIccLoout1Off = Nothing
    Set mpldVldLoout1 = Nothing
    Set mpldVldFqm3 = Nothing
    Set mpldIccLoout2On = Nothing
    Set mpldIccLoin2On = Nothing
    Set mpldVldLoin2 = Nothing
    Set mpldVldInFqm3 = Nothing
    Set mpldVldLoout2 = Nothing
    Set mpldIccFqm3Off = Nothing
    Set mpldIccCasSplitOff = Nothing
    Set mpldIccCasCombOff = Nothing
    Set mpldIccFqm2Off = Nothing
    
    For liCount = 0 To 12
        mpldaTxloIccResults(liCount) = IfxMath.Calc.PinListData.Create(-999, lsPins)
    Next liCount
    For liCount = 0 To 7
        mpldaTxloVldResults(liCount) = IfxMath.Calc.PinListData.Create(-999, "AMUX1P_DIFF")
    Next liCount
    
    msPinsCurrCons = lsPins
    msVdd1v8Tx = lsVdd1v8Pin
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function performs the TXLO VLD offset measurements
'''@details
'''Measurements are executed inside a Dopp pattern using the DiffMeter on AMUX1. The pattern finishes with a SoftReset
''' The measured values are read back at the end of the pattern and stored in module level variables for later
''' offset compensation and datalogging.
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsVldOffsets()
    Dim lsPatNameExt As String
    Dim liIdx As Long
    
    'DOPP-PATTERN
    '''-# VLD offset measurement pattern at Vdd-5% using Diffmeter at AMUX to measure offsets for compensation
    lsPatNameExt = "LoVldOffset"
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)

    ''' -# Disable the DFX gate protection
    Call Ctrx.DisableDfxGateProtection
    ''' -# Enable TX central biasing
    Call TxCentralBias(OptTxCentralBias_En)
    ''' -# Enable the central TXLO biasing
    Call TxCtrlLoBias(OptLoBias_En)
    ''' -# Global Amux configured to connect PAD 1 Differential and disable PAD 2
    Call TxConfigGlobalAmux(OptAmuxPADConfType_Pad1En, OptAmuxConfType_Differential, _
                            OptAmuxPADConfType_Pad2Disable, OptAmuxConfType_Disconnect)

    ''' -# TXLO VLDs offset measurements
    Call TxgeTxloMeasVldOffsets
    ''' -# Enable back the DFX protection
    Call Ctrx.EnableDfxGateProtection
    ''' -# Soft-Reset and pattern stop
    Call Ctrx.SoftReset(peUpdateMethod:=ForceUpdate, pbAlarmToggleRequired:=True)
    Call Ctrx.DoppStopPattern
    
    ''' -# Read back Measured VLDs:
    mpldVldInFqm2Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldOutFqm2Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldInFqm3Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldOutFqm3Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldLoin1Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldLoin2Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldLoout1Offset = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldLoout2Offset = DoppRead("AMUX1P_DIFF", Diffmeter)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Perform the TXLO VLDs offset measurements
'''@details
''' This routine must be executed inside a Dopp block
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloMeasVldOffsets()
    ''' Enable the VLDs
    Call TxgeTxloCurConsEnableVlds
    ''' Select the FQM2 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x1, WAIT_5MS) 'Wait time from STO
    ''' Select the FQM2 VLD OUT channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x2, WAIT_5MS) 'Wait time from STO
    ''' Select the FQM3 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x3, WAIT_5MS) 'Wait time from STO
    ''' Select the FQM3 VLD OUT channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x4, WAIT_5MS) 'Wait time from STO
    ''' Select the LOIN1 VLD channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x14, WAIT_1MS)
    ''' Select the LOIN2 VLD channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x15, WAIT_1MS)
    ''' Select the LOOUT1 VLD channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0xD, WAIT_1MS)
    ''' Select the LOOUT2 VLD channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0xE, WAIT_1MS)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Routine enables the TXLO VLDs for the offset measurements
'''@details
''' This routine must be executed inside a Dopp block
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsEnableVlds()
    ''' Enable the FQM3 VLDs
    With IfxRegMap.NewRegTxTxloFqm3Cf
        .bfEnVldIn = 1
        .bfEnVldOut = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' Enable the FQM2 VLDs
    With IfxRegMap.NewRegTxTxloFqm2Cf
        .bfEnVldIn = 1
        .bfEnVldOut = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' Enable the LOIN1 and LOIN2 VLDs
    With IfxRegMap.NewRegTxTxloLoinCf
        .bfEnLoin1Vld = 1
        .bfEnLoin2Vld = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' Enable the LOOUT1 and LOOUT2 VLDs
    With IfxRegMap.NewRegTxTxloLooutCf
        .bfEnLoout1Vld = 1
        .bfEnLoout2Vld = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Run the Dopp pattern to measure the Vdd supplies uncompensated.
'''@details Measures the Vdd voltage for VDD0V9RF, VDD0V9PA and VDD1V8PA DCVIs using the AMUX. For the VddPA values,
''' one Tx channel is enabled, default TXCHANNEL1.
''' In total, 11 voltage measurements are made, all using AMUX1P_DIFF.
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsPreconditioning()
    
    Dim lsPatNameExt As String
    Dim liTxChannel As Long
    Dim loPreOpCapt As FirmwareVariable
    Dim loLowPowCapt As FirmwareVariable
    Dim loOpCapt As FirmwareVariable
    Dim loFwCmdExitPreOperation As FwCmdExitPreOperation
    Dim loFwCmdGotoOperation As FwCmdGotoOperation
        
    'DOPP-PATTERN
    '''-# Preconditioning measurement pattern at Vdd-5% using Diffmeter at AMUX to measure offsets for compensation
    lsPatNameExt = "Pre"
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' -# Program the DUT to the operation-mode, monitor the TX TOP states and configure the DPLL operating frequency
    Call DutPreopToOperation( _
        loFwCmdExitPreOperation, loFwCmdGotoOperation, False, , , True, loPreOpCapt, loLowPowCapt, loOpCapt)
    ''' -# Disable the DFX gate protection
    Call Ctrx.DisableDfxGateProtection
    ''' -# Monitor the chip temperature before starting measurement
    Call DtsMonitoringNoPattern(moFwCmdGetTempPre)
    ''' -# Disable the RX-subsystem
    Call RxRxChansDisabling(WAIT_100US)
    ''' -# Disable all Txrf phase shifters
    Call TxrfPsDisable
    ''' -# Disable the local biasing for each TXRF channel
    For liTxChannel = TXCHANNEL1 To TXCHANNEL8
        With CallByName(IfxRegMap, "NewRegTxTxrf" & CStr(liTxChannel) & "BiasCf", VbMethod)
            .bfEn = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    Next liTxChannel

    ''' -# Global Amux configured to connect PAD 1 Differential and disable PAD 2
    Call TxConfigGlobalAmux(OptAmuxPADConfType_Pad1En, OptAmuxConfType_Differential, _
                            OptAmuxPADConfType_Pad2Disable, OptAmuxConfType_Disconnect)
    
    ''' Step 13 Txlo Supply voltage measurements - MIN supply levels
    Call TxgeTxloMeasureSupplyLevels
    ''' Recover measured values, calculate compensation values and apply immediately to the DCVIs
    Dopp.PatternPause
    Call TxgeTxloCurConsReadVddPre
    ''' Measure the compensated supply voltages
    Call TxgeTxloMeasureSupplyLevels
    ''' Pattern stop for Preconditioning results readback and Supply compensation calculations
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN

    ''' Read the post compensated supply voltages. (0) = 1V8, (1) = 1V0
    mpldaTxloVddPost(0) = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldaTxloVddPost(1) = DoppRead("AMUX1P_DIFF", Diffmeter)

    msdPreOpDtct = loPreOpCapt.Value
    msdLowPowDtct = loLowPowCapt.Value
    msdOpDtct = loOpCapt.Value
    msdExitPreOpErr = loFwCmdExitPreOperation.Errorcode
    msdGotoOpErr = loFwCmdGotoOperation.Errorcode

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Routine to measure TXLO Vdd1V8 and TXLO LOOUT2 Vdd1V0 supplies
'''@details
''' This subroutine must be called from within a Dopp block. Supply voltages are routed out to the AMUX1 pins.
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloMeasureSupplyLevels()
    
    ''' Txlo Vdd1v8 Supply voltage measurement
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0xC, WAIT_1MS)
    
    ''' Txlo LoOut2 Vdd1v0 Supply voltage measurement
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x1C, WAIT_1MS)
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Run the Dopp pattern to measure the Vdd supplies for TXLO compensation.
'''@details
''' Measures the Vdd voltage for VDD1V8 and VDD1V0 DCVIs using the AMUX. In total, 2 voltage measurements are made,
''' all using AMUX1P_DIFF. The required compensation is calculated and then applied.
''' The function exits with the Dopp pattern paused
'''
'''@log{2025-11-11, Channon Andrew, initial implementation\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsReadVddPre()
    
    Dim liIdx As Long
    Dim lvSite As Variant
    Dim lsPin As String
    Dim lpldVddMeas As New PinListData
    Dim lpldVddCorr As New PinListData
    Dim lpldVdd1V0 As New PinListData
    Dim lpldVdd1V8 As New PinListData
    Dim ls1V8DcviPin As String
    Dim lwTxloVddCorrMin As New DSPWave
        
    ''' Initialize all variables for measurement and subsequent calculation
    lsPin = "AMUX1P_DIFF"
    Call lwTxloVddCorrMin.CreateConstant(0, 2, DspDouble)
    
    ' Initialize variables for calculations
    lpldVdd1V0 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V0 * VDDMIN_PERCENT, lsPin)
    lpldVdd1V8 = IfxMath.Calc.PinListData.Create(VOLTAGE_1V8 * VDDMIN_PERCENT, lsPin)
    ' Loop for readback of uncompensated VDD measurements with calculation of compensation values
    For liIdx = 0 To 1
        lpldVddMeas = DoppRead(lsPin, Diffmeter)
        ' Calculate the correction factor based on the value measured
        ' then check against extreme voltage compensation levels in case something goes wrong
        If (liIdx Mod 2) = 0 Then
            ' even values of loop index for 1V8 measured values
            lpldVddCorr = lpldVdd1V8.Math.Add(lpldVdd1V8).Subtract(lpldVddMeas)
            lpldVddCorr = IfxMath.Calc.PinListData.Clip(lpldVddCorr, VDD1V8_THR_MIN, VDD1V8_THR_MAX)
        Else
            ' odd values of loop index for 1V0 measured values
            lpldVddCorr = lpldVdd1V0.Math.Add(lpldVdd1V0).Subtract(lpldVddMeas) 'TxCh 5-8
            lpldVddCorr = IfxMath.Calc.PinListData.Clip(lpldVddCorr, VDD1V0_THR_MIN, VDD1V0_THR_MAX)
        End If
        '''-# Convert the PinlistData values to DspWave, the format used to apply the levels
        For Each lvSite In TheExec.Sites
            lwTxloVddCorrMin(lvSite).Element(liIdx) = lpldVddCorr.Pins(lsPin)(lvSite)
        Next lvSite
    Next liIdx
    
    ''' Apply compensated supply voltage levels
    If (msHwSetTpConfig = "FE") Then
        ls1V8DcviPin = "VDD1V8TX_DCVI"
    Else
        ls1V8DcviPin = "VDD1V8_DCVI"
    End If
    
    For Each Site In TheExec.Sites
        Call IfxHdw.Dcvi.SetVoltage(ls1V8DcviPin, lwTxloVddCorrMin(Site).Element(0))
        Call IfxHdw.Dcvi.SetVoltage("VDD1V0TX_MDCVI", lwTxloVddCorrMin(Site).Element(1))
    Next Site

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief This function configures the DCVI instruments for the main measurement pattern
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsInstrSetup()
    Dim lsDcviPins As String

    '''-# Setup voltage supply DCVIs to measure current
    If (msHwSetTpConfig = "FE") Then
        lsDcviPins = "VDD1V8TX_DCVI, VDD1V0TX_MDCVI"
    Else
        lsDcviPins = "VDD1V8_DCVI, VDD1V0TX_MDCVI"
    End If
    TheHdw.Dcvi.Pins(lsDcviPins).Meter.Mode = tlDCVIMeterCurrent
    TheHdw.Dcvi.Pins(lsDcviPins).Meter.HardwareAverage = HW_AVERAGE_64
    TheHdw.Dcvi.Pins(lsDcviPins).Capture.SampleRate = 1 * MHz
    Call IfxHdw.Dcvi.BleederResistorOff(lsDcviPins)
    
    Call IfxHdw.Dcvi.SetCurrent("VDD1V0TX_MDCVI", 400 * mA, 400 * mA)
    
    '''-# Update to the appropriate wait time for the DCVI pattern strobes
    mdDcviWait = IfxHdw.Dcvi.GetStrobeWait("VDD1V0TX_MDCVI")
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Measure the ICC currents from the Tx Supplies and measure the compensated supply voltages
'''@details
''' Run the pattern to measure the supply currents for VDD1V8TX and VDD1V0TX at minimum and maximum levels. Here we
''' also measure the compensated supply voltages using the Diffmeter via the AMUX.
''' Afterwards, read back the measured values and store for later processing and datalogging
''' Disconnect the DiffMeters after reading back the measured values to clear the setup for the next step.
'''
'''@log{2025-09-12, Channon Andrew, Initial Version}
'''@log{2025-09-12, Channon Andrew, Adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsMeas()
    Dim lsPatNameExt As String
    Dim liIdx As Long
    Dim lslSweepIdx As New SiteLong
    Dim lpldLoin1PathTrim As New PinListData
    Dim lsaIdacCodeSame() As String
    
    lsaIdacCodeSame = Split(LO1TST_BUF_DAC_CODE, ",")
    
    'DOPP-PATTERN
    lsPatNameExt = "Icc" & msPatHw
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName(psNameExtension:=lsPatNameExt), pbReadFailBitCnt:=False)
    ''' Steps 15, 16, 17 for Standalone and LO12 STA buffers
    Call TxgeTxloCurConsMeasStandalone
    
    ''' Step 18. LOIN1 Path. ICC and VLD measurements
    ''' - Enable the LOIN1 buffer and VLD
    Call TxLoLoIn12(OptLoIn1_En, OptLoIn2_Disable)
    ''' - Enable the LO1 TST buffer
    Call TxgeLoTstBuf1En(OptLoTstBuf1_En)
    ''' - Enable the LOOUT1 buffer
    Call TxLoLoOut(OptLoOut1_En, OptLoOut2_Disable, OptLoOutBuf_En)
    ''' - Select the LOIN1 VLD channel from the TXLO AMUX
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = OptTxTxraigroupTxloSenseConf_AmuxSel_0x14
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' - Sweep Lo1_Tst and measure voltage
    lpldLoin1PathTrim = TxgeTxloTstBufSweep
    ''' - Apply found target trim code and measure current
    lslSweepIdx = 0
    lslSweepIdx = lpldLoin1PathTrim.Pins("AMUX1P_DIFF").Compare(EqualTo, TRIMM_ERROR).If _
        (TX_TRIM_ERR_0, lpldLoin1PathTrim.Pins("AMUX1P_DIFF"))
    Call IfxMath.Calc.SiteLong.ReplaceIf(lslSweepIdx, lslSweepIdx.Compare(GreaterThanOrEqualTo, 32), 31)
    ''' - Store Lo1 path trim values for datalogging
    For Each Site In TheExec.Sites
        mslTstBufSweepValRead(Site) = CLng(lsaIdacCodeSame(lslSweepIdx(Site)))
        msdTstBufSweepValReadV(Site) = lpldLoin1PathTrim.Pins("AMUX1P_DIFF_VAL").Value
    Next Site
    ''' - Write found trim target
    With IfxRegMap.NewRegTxTxloTestStaCtrl
        .bfCtrlTestbuf1ms = mslTstBufSweepValRead
        Call Ctrx.WriteVolatileRegister(.Self)
    End With
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Select the LOOUT1 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0xD, WAIT_1MS)
    ''' - Select the FQM3 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x3, WAIT_5MS) 'Wait time from STO
    ''' - Disable the LOIN1 buffer and VLD
    Call TxLoLoIn12(OptLoIn1_Disable, OptLoIn2_Disable)
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the LO1 TST buffer
    Call TxgeLoTstBuf1En(OptLoTstBuf1_Disable)
    ''' - Disable the LOOUT1 buffer
    Call TxLoLoOut(OptLoOut1_Disable, OptLoOut2_Disable)
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    
    ''' Steps 19, 20 for LOIN2 path and remaining blocks
    Call TxgeTxloCurConsMeasLoin2
    
    ''' -# Monitor the DTS after the measurements and before resetting the DUT
    Call DtsMonitoringNoPattern(moFwCmdGetTempPost)
    ''' -# Step 22. Enable back the DFX protection
    Call Ctrx.EnableDfxGateProtection
    ''' -# Step 23. Soft-Reset and pattern stop
    Call Ctrx.SoftReset(peUpdateMethod:=ForceUpdate, pbAlarmToggleRequired:=True)
    Call Ctrx.DoppStopPattern
    'END DOPP-PATTERN

    '''-# Read measured voltages and currents after the LOIN1 Path Trimming
    Call TxgeTxloReadPostLoin1TrimMeasValues
    
    ''' Disconnect DIFFMETERs
    Call TxAmuxDiffmeterSetup(DCVI_DIFF_DISC)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Txlo Lo1 TstBuffer sweep and post processing
'''@details
''' Perform a register sweep for Lo1 Test Buffer for a defined set of 32 IDAC code values
''' Read back all measured values and find the value closest to the given target value
'''
'''@return [As PinListData] PinListData object containing the Trimming value and the sweep position where found
'''
'''@log{2024-03-19, Olia Svyrydova, Initial Version}
'''@log{2024-11-12, Neul Roland, move SPI/JTAG support to library\, @JIRA{CTRXTE-5202}}
'''@log{2025-02-10, Channon Andrew, Remove not needed register readback\, @JIRA{CTRXTE-5553}}
'''@log{2025-09-12, Channon Andrew, Adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function TxgeTxloTstBufSweep() As PinListData
    Dim liLoopCounter As Long
    Dim liLo1CodeStep As Long
    Dim lsaIdacCodeSame() As String
    Dim lpldaMeasValues() As New PinListData
    Dim lpldResultData As New PinListData

    '''-# Add Pins to the result Object
    lpldResultData.AddPin("AMUX1P_DIFF").Value = TRIMM_ERROR
    lpldResultData.AddPin("AMUX1P_DIFF_VAL").Value = -99999#
    
    '''-# Logging start the sweep trimming
    Call Logger.Dbug("Logging Start Sweep-Trimming")
    
    '''-# Do the sweeping
    lsaIdacCodeSame = Split(LO1TST_BUF_DAC_CODE, ",")
    For liLo1CodeStep = LBound(lsaIdacCodeSame) To UBound(lsaIdacCodeSame)
        With IfxRegMap.NewRegTxTxloTestStaCtrl
            .bfCtrlTestbuf1 = CLng(lsaIdacCodeSame(liLo1CodeStep))
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        Call Dopp.Settle(WAIT_1MS)
        Call Dopp.DCDiffMeter("AMUX1P_DIFF").Strobe(mdPaDiffMeterWait)
    Next liLo1CodeStep
    
    '''-# Pause Pattern to read all measured values before the LOIN1 sweep and then find the trim value from sweep
    Call Ctrx.DoppPausePattern
    Call TxgeTxloReadPreLoin1TrimMeasValues
    
    '''-# Loop for readout of the strobed values and evaluate the closest result to the given target value
    ReDim lpldaMeasValues(0 To LO1TST_CNT)
    For liLoopCounter = 0 To LO1TST_CNT
        lpldaMeasValues(liLoopCounter) = DoppRead("AMUX1P_DIFF", Diffmeter)
        ' apply offset correction
        lpldaMeasValues(liLoopCounter) = lpldaMeasValues(liLoopCounter).Math.Subtract(mpldVldLoin1Offset)
        Call TxgeTxloProcessLoout1TstBufSweep(lpldResultData, lpldaMeasValues(liLoopCounter), liLoopCounter)
    Next liLoopCounter
    
    '''-# Check for valid value. If not assign TRIMM_ERROR (-99)
    For Each Site In TheExec.Sites
        If ((lpldResultData.Pins("AMUX1P_DIFF").Value < TST_BUF_START_VAL) Or _
            (lpldResultData.Pins("AMUX1P_DIFF").Value > (TST_BUF_END_VAL))) Then
            
            lpldResultData.Pins("AMUX1P_DIFF").Value = TRIMM_ERROR
        End If
    Next Site
      
    '''-# Logging the end of the sweep
    Call Logger.Dbug("End Sweep")
    Call Logger.Dbug("")
    
    '''-# return trimming result and corresponding sweep index value
    Set TxgeTxloTstBufSweep = lpldResultData
    
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief TxgeTxloProcessLoout1TstBufSweep find closest to the target value from sweep
'''
'''@param[in/out] ppldValRes result
'''@param[in/out] ppldMeasValues measured value
'''@param[in/out] piLoopCounter loop counter
'''
'''@log{2024-03-15, Olia Svyrydova, Initial Version}
'''@log{2024-07-03, Olia Svyrydova, Update for 8191B \, @JIRA{CTRXTE-4660}}
'''@log{2024-07-26, vommi, Update if condition 8191B \, @JIRA{CTRXTE-4841}}
'''@log{2025-09-12, Channon Andrew, Adapted for 8188A\, @JIRA{RSIPPTE-109}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloProcessLoout1TstBufSweep( _
    ByRef ppldValRes As PinListData, _
    ByRef ppldMeasValues As PinListData, _
    ByVal piLoopCounter As Long)
    
    Dim lsdMeasValue1  As New SiteDouble
       
    lsdMeasValue1 = ppldMeasValues.Pins("AMUX1P_DIFF")
    If TheExec.TesterMode = testModeOnline Then
        For Each Site In TheExec.Sites
            '''-# Take the closest
            '''-# Check if the actual value is the closest
            If Abs((lsdMeasValue1(Site) - VLD_TARGET_VALUE) < _
                Abs(ppldValRes.Pins("AMUX1P_DIFF_VAL").Value - VLD_TARGET_VALUE)) Then
                    
                '''-# Assign result as new final result.
                ppldValRes.Pins("AMUX1P_DIFF_VAL").Value = lsdMeasValue1(Site)
                '''-# Store the number of the step as final register result.
                ppldValRes.Pins("AMUX1P_DIFF").Value = piLoopCounter + TST_BUF_START_VAL
            End If
            
        Next Site
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Execute first part of TxgeTxloCurConsMeas measurement pattern
'''@details
''' Runs the Standalone measurement steps and the LO1/2 STA buffer measurements
''' Must be executed inside a Dopp block
'''
'''@log{2025-11-13, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsMeasStandalone()
    
    ''' Step 15. ICC Measurements Standalone
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the TXLO TXMON
    With IfxRegMap.NewRegTxTxloMonCf0
        .bfEn = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the TXLO distribution (SSD + ISD + Split blocks)
    With IfxRegMap.NewRegTxTxloLodistCf
        .bfEnSsdTxlo = 0
        .bfEnSsdRxlo = 0
        .bfEnAmpTxLeft = 0
        .bfEnAmpTxRight = 0
        .bfEnAmpTxRx = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' - Keep the TXLO local biasing and VLDs biasing enabled while disabling the biasing of the LO distribution _
        and BBDs units
    With IfxRegMap.NewRegTxTxloBiasCf
        .bfLoEnBias = 1
        .bfEnVldBias = 1
        .bfEnLoBbd = 0
        .bfEnLodistBias = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    
    ''' Step 16. VLD Measurements Standalone
    ''' - Configure the FQM2 block and its VLDs
    With IfxRegMap.NewRegTxTxloFqm2Cf
        .bfEnBias = 1
        .bfEnMirror = 1
        .bfEnBiasCore = 1
        .bfEnVldIn = 1
        .bfEnVldOut = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' - Configure the FQM3 block and its VLDs
    With IfxRegMap.NewRegTxTxloFqm3Cf
        .bfEnBias = 1
        .bfEnMirror = 1
        .bfEnBiasCore = 1
        .bfEnVldIn = 1
        .bfEnVldOut = 1
        .bfEnBiasBuf2 = 1
        .bfEnBiasBuf3 = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' - Select the FQM2 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x1, WAIT_5MS) 'Wait time from STO
    ''' - Select the FQM2 VLD OUT channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x2, WAIT_5MS) 'Wait time from STO
    ''' - Select the FQM3 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x3, WAIT_5MS) 'Wait time from STO
    ''' - Select the FQM3 VLD OUT channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x4, WAIT_5MS) 'Wait time from STO
    
    ''' Step 17. LO1, LO2 STA buffers. ICC Measurements
    ''' - Disable the LO1 stand-alone buffer
    Call TxloStaBufConfig(pbSta1Enable:=False, pbSta2Enable:=False)
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Enable the LO2 stand-alone buffer
    Call TxloStaBufConfig(pbSta1Enable:=False, pbSta2Enable:=True)
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the LO2 stand-alone buffer
    Call TxloStaBufConfig(pbSta1Enable:=False, pbSta2Enable:=False)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Execute the part of TxgeTxloCurConsMeas measurement pattern after the LOIN1 Path trimming sweep
'''@details
''' Runs the LOIN2 path testing and the remaining blocks
''' Must be executed inside a Dopp block
'''
'''@log{2025-11-13, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsMeasLoin2()
    
    ''' Step 19. LOIN2 Path. ICC and VLD measurements
    ''' - Enable the LOOUT2 buffer
    Call TxLoLoOut(OptLoOut1_Disable, OptLoOut2_En, OptLoOutBuf_En)
    ''' - Program the LOOUT2 buffer stage 1
    With IfxRegMap.NewRegTxTxloLoout2Ctrl
        .bfCtrlIdacStage1 = 112
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    ''' - Enable the LO2 TST buffer
    Call TxgeLoTstBuf2En(OptLoTstBuf2_En)
    ''' - Write found trim target from LOIN1 path test
    With IfxRegMap.NewRegTxTxloTestStaCtrl
        .bfCtrlTestbuf2ms = mslTstBufSweepValRead
        Call Ctrx.WriteVolatileRegister(.Self)
    End With
    Call Dopp.Wait(WAIT_1MS)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Enable the LOIN2 buffer and VLD
    Call TxLoLoIn12(OptLoIn1_Disable, OptLoIn2_En)
    Call Dopp.Wait(WAIT_1MS)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Switch to the CASSPLIT2 and CASCOMB2 buffers in order to close the LOIN2 path
    Call TxloCasCombConfig(pbEnable:=True, peOptSelBuf:=OptSelBuf_En)
    Call TxloCassplitConfig(pbEnable:=True, peOptSelBuf:=OptSelBuf_En)
    ''' - Select the LOIN2 VLD channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x15, WAIT_1MS)
    ''' - Select the FQM3 VLD IN channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0x3, WAIT_5MS) 'Wait time from STO
    ''' - Select the LOOUT2 VLD channel from the TXLO AMUX
    Call TxgeTxloAmuxStrobe(OptTxTxraigroupTxloSenseConf_AmuxSel_0xE, WAIT_1MS)
    
    ''' Step 20. ICC Measurements - Remaining Blocks
    ''' - Disable the FQM3 block
    Call TxFqm3Disable
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the CASSPLIT buffers
    Call TxloCassplitConfig(pbEnable:=False)
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the CASCOMB buffers
    Call TxloCasCombConfig(pbEnable:=False)
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)
    ''' - Disable the FQM2 block
    Call TxFqm2Disable
    Call Dopp.Wait(WAIT_500US)
    Call Dopp.Dcvi.Pins(msPinsCurrCons).Meter.Strobe(mdDcviWait)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief For TxloCurrCons, read back all instrument measured values before the LOIN1 path trimming
'''
'''@log{2025-11-13, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloReadPreLoin1TrimMeasValues()

    ''' 15. TXLO ICCs measurements - Standalone operation
    mpldIccStandalone = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccTxmonOff = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccLoDistOff = DoppRead(msPinsCurrCons, Dcvi)
    ''' 16. FQM2/3 VLDs measurements - Standalone operation
    mpldVldInFqm2Standalone = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldOutFqm2Standalone = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldInFqm3Standalone = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldOutFqm3Standalone = DoppRead("AMUX1P_DIFF", Diffmeter)
    ''' 17. TXLO ICC measurements - LO1/2 STA buffers
    mpldIccLo1StaOff = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccLo2StaOn = DoppRead(msPinsCurrCons, Dcvi)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief For TxloCurrCons, read back all instrument measured values after the LOIN1 path trimming
'''
'''@log{2025-11-13, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloReadPostLoin1TrimMeasValues()

    ''' 18. TXLO ICC and VLDs measurements - LOIN1 path testing
    mpldIccLo1Path = DoppRead(msPinsCurrCons, Dcvi)
    mpldVldLoout1 = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldFqm3 = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldIccLoin1Off = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccLoout1Off = DoppRead(msPinsCurrCons, Dcvi)
    ''' 19. LOIN2 Path. ICC and VLD measurements
    mpldIccLoout2On = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccLoin2On = DoppRead(msPinsCurrCons, Dcvi)
    mpldVldLoin2 = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldInFqm3 = DoppRead("AMUX1P_DIFF", Diffmeter)
    mpldVldLoout2 = DoppRead("AMUX1P_DIFF", Diffmeter)
    ''' Step 20. ICC Measurements - Remaining Blocks
    mpldIccFqm3Off = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccCasSplitOff = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccCasCombOff = DoppRead(msPinsCurrCons, Dcvi)
    mpldIccFqm2Off = DoppRead(msPinsCurrCons, Dcvi)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Calculate the TX current consumption values from all measured supply values
'''@details
'''Calculate the delta current consumptions
'''
'''@log{2025-11-13, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsCalc()

    mpldaTxloIccResults(0) = mpldIccStandalone
    mpldaTxloIccResults(1) = TxgeTxSysDeltaCurr(mpldIccStandalone, mpldIccTxmonOff)
    mpldaTxloIccResults(2) = TxgeTxSysDeltaCurr(mpldIccTxmonOff, mpldIccLoDistOff)
    mpldaTxloIccResults(3) = TxgeTxSysDeltaCurr(mpldIccLoDistOff, mpldIccLo1StaOff)
    mpldaTxloIccResults(4) = TxgeTxSysDeltaCurr(mpldIccLo2StaOn, mpldIccLo1StaOff)
    mpldaTxloIccResults(5) = TxgeTxSysDeltaCurr(mpldIccLo1Path, mpldIccLoin1Off)
    mpldaTxloIccResults(6) = TxgeTxSysDeltaCurr(mpldIccLoin1Off, mpldIccLoout1Off)
    mpldaTxloIccResults(7) = TxgeTxSysDeltaCurr(mpldIccLoout2On, mpldIccLoout1Off)
    mpldaTxloIccResults(8) = TxgeTxSysDeltaCurr(mpldIccLoin2On, mpldIccLoout2On)
    mpldaTxloIccResults(9) = TxgeTxSysDeltaCurr(mpldIccLoout2On, mpldIccFqm3Off)
    mpldaTxloIccResults(10) = TxgeTxSysDeltaCurr(mpldIccFqm3Off, mpldIccCasSplitOff)
    mpldaTxloIccResults(11) = TxgeTxSysDeltaCurr(mpldIccCasSplitOff, mpldIccCasCombOff)
    mpldaTxloIccResults(12) = TxgeTxSysDeltaCurr(mpldIccCasCombOff, mpldIccFqm2Off)

    mpldaTxloVldResults(0) = mpldVldInFqm2Standalone.Math.Subtract(mpldVldInFqm2Offset)
    mpldaTxloVldResults(1) = mpldVldOutFqm2Standalone.Math.Subtract(mpldVldOutFqm2Offset)
    mpldaTxloVldResults(2) = mpldVldInFqm3Standalone.Math.Subtract(mpldVldInFqm3Offset)
    mpldaTxloVldResults(3) = mpldVldOutFqm3Standalone.Math.Subtract(mpldVldOutFqm3Offset)
    mpldaTxloVldResults(4) = mpldVldLoout1.Math.Subtract(mpldVldLoout1Offset)
    mpldaTxloVldResults(5) = mpldVldFqm3.Math.Subtract(mpldVldInFqm3Offset)
    mpldaTxloVldResults(6) = mpldVldLoin2.Math.Subtract(mpldVldLoin2Offset)
    mpldaTxloVldResults(7) = mpldVldLoout2.Math.Subtract(mpldVldLoout2Offset)
    mpldaTxloVldResults(8) = mpldVldInFqm3.Math.Subtract(mpldVldInFqm3Offset)

End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Datalog all results for the Txge Txlo Current Consumption tests
'''
'''@log{2025-11-13, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloCurConsDatalog()
    Dim lsVdd1v8Pin1 As String
    Dim lsVdd1v0Pin2 As String
    Dim liCount As Long
    
    '''- Define the supply pins for datalogging current
    If (msHwSetTpConfig = "FE") Then
        lsVdd1v8Pin1 = "VDD1V8TX_DCVI"
    Else
        lsVdd1v8Pin1 = "VDD1V8_DCVI"
    End If
    lsVdd1v0Pin2 = "VDD1V0TX_MDCVI"
    
    ''' TX pre-conditioning status monitoring values
    Call IfxEnv.Datalog(msdPreOpDtct)
    Call IfxEnv.Datalog(msdLowPowDtct)
    Call IfxEnv.Datalog(msdOpDtct)
    Call IfxEnv.Datalog(msdExitPreOpErr)
    Call IfxEnv.Datalog(msdGotoOpErr)
    ''' Chip temperature readouts before and after the measurements
    Call IfxEnv.Datalog(moFwCmdGetTempPre.Temperature(DTS_TEMP1))
    Call IfxEnv.Datalog(moFwCmdGetTempPre.Temperature(DTS_TEMP2))
    Call IfxEnv.Datalog(moFwCmdGetTempPost.Temperature(DTS_TEMP1))
    Call IfxEnv.Datalog(moFwCmdGetTempPost.Temperature(DTS_TEMP2))

    ''' Datalog compensated supply levels
    Call IfxEnv.Datalog(mpldaTxloVddPost(0))
    Call IfxEnv.Datalog(mpldaTxloVddPost(1))
    
    ''' Datalog ICCs
    For liCount = 0 To 12
        Call TxgeDcCurrDatalog(mpldaTxloIccResults(liCount), lsVdd1v8Pin1, lsVdd1v0Pin2)
    Next liCount
    ''' Datalog VLD offsets
    Call IfxEnv.Datalog(mpldVldInFqm2Offset)
    Call IfxEnv.Datalog(mpldVldOutFqm2Offset)
    Call IfxEnv.Datalog(mpldVldInFqm3Offset)
    Call IfxEnv.Datalog(mpldVldOutFqm3Offset)
    Call IfxEnv.Datalog(mpldVldLoin1Offset)
    Call IfxEnv.Datalog(mpldVldLoin2Offset)
    Call IfxEnv.Datalog(mpldVldLoout1Offset)
    Call IfxEnv.Datalog(mpldVldLoout2Offset)
    ''' Datalog VLDs
    For liCount = 0 To 8
        Call IfxEnv.Datalog(mpldaTxloVldResults(liCount))
        If liCount = 3 Then
            ''' insert the datalogging for the LOIN1 Path trimming
            Call IfxEnv.Datalog(msdTstBufSweepValReadV)
            Call IfxEnv.Datalog(mslTstBufSweepValRead)
        End If
    Next liCount

End Sub

' HELPERS

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Helper function measures the 1V8 and 1V0 supply voltages via the AMUX
'''@details
''' This subroutine must be used inside a Dopp block and it must be inserted inside a For Loop for the selection
''' of one of the 8 TXRF channels. This routine is primarily intended to support the Txge Current Consumption
''' measurements where supply voltage compensation is required for each channel.
'''
'''@param[in] piIdx The loop index value between 1 and 8 for the selected Txrf Channel
'''@param[in] psMeasPin The Diffmeter pin name for the voltage measurement. Differs for Txrf channels 1-4 and 5-8
'''@param[in] pbTxCh1toTxCh4 Boolean switch parameter to identify if selected Txrf channel is in Ch1-4 or Ch5-8
'''@param[in] pdWait1V8 Optional override capability for defining settling time for 1V8 measurement
'''@param[in] pdWait1V0 Optional override capability for defining settling time for 1V0 measurement
'''
'''@log{2025-09-29, Channon Andrew, Initial Version}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function TxgeMeasureAmuxSupplyVoltages( _
    ByVal piIdx As Long, _
    ByVal psMeasPin As String, _
    ByVal pbTxCh1toTxCh4 As Boolean, _
    Optional ByVal pdWait1V8 As Double = WAIT_1MS, _
    Optional ByVal pdWait1V0 As Double = WAIT_2MS)
    
    '''-# Measure the VDD1V8 voltage at the AMUX
    Call TxTxrfSingleAmuxSel(piIdx, pdVoltage:=VOLTAGE_1V8)
    Call Dopp.Settle(pdWait1V8)
    Call Dopp.DCDiffMeter(psMeasPin).Strobe(mdPaDiffMeterWait)
    '''-# Measure the VDD1V0 voltage at the AMUX
    Call TxTxrfSingleAmuxSel(piIdx, pdVoltage:=VOLTAGE_1V0)
    Call Dopp.Settle(pdWait1V0)
    Call Dopp.DCDiffMeter(psMeasPin).Strobe(mdPaDiffMeterWait)
    '''-# Reset the AMUX then ramp down the IDAC and disable the phase shifter and PA of the current TXRF block
    Call TxTxrfSingleAmuxSel(piIdx, pbReset:=True)
    Call TxrfRampPaDac(TxGetTxChanPair(piIdx), False, pbTxCh1toTxCh4, Not pbTxCh1toTxCh4)
    Call TxDisablePAChannels(piIdx, pbPaCfEnVld:=True)
    Call TxrfPsSingleDisable(piIdx)

End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Helper routine to set up the TXLO AMUX, apply desired settling and strobe measurement with the DiffMeter
'''@details
''' Must be executed from inside a Dopp block
'''
'''@param[in] peAmuxSel The TXLO AMUX selection
'''@param[in] pdWait The settling time to be applied
'''
'''@log{2025-11-11, Channon Andrew, Initial Version\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxloAmuxStrobe( _
    ByVal peAmuxSel As OptTxTxraigroupTxloSenseConf_AmuxSel, _
    ByVal pdWait As Double)
    
    With IfxRegMap.NewRegTxTxloSenseCf
        .bfAmuxSel = peAmuxSel
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Wait(pdWait)
    Call Dopp.DCDiffMeter.Pins("AMUX1P_DIFF").Strobe(mdPaDiffMeterWait)
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Helper routine to Enable or Disable the LO1 TST buffer
'''
'''@param[in] poLoTstBuf1 Selection for Enable / Disable
'''
'''@log{2024-02-07, Olia Svyrydova, Initial Version}
'''@log{2025-11-12, Channon Andrew, ported from 8191\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeLoTstBuf1En( _
    Optional poLoTstBuf1 As OptLoTstBuf)

    If poLoTstBuf1 Then
        With IfxRegMap.NewRegTxTxloTestStaCf
            .bfEnTestbuf1 = 1
            .bfEnStabuf1Filterbypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
            Call Dopp.Wait(SR_TX_BIAS_EN_FILTERBYPASS_WAIT)
            .bfEnStabuf1Filterbypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    Else
        With IfxRegMap.NewRegTxTxloTestStaCf
            .bfEnTestbuf1 = 0
            .bfEnStabuf1Filterbypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
            Call Dopp.Wait(SR_TX_BIAS_EN_FILTERBYPASS_WAIT)
            .bfEnStabuf1Filterbypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    Call Dopp.Wait(SR_TXLO_LOIN_CONF_EN_WAIT)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Helper routine to Enable or Disable the LO2 TST buffer
'''
'''@param[in] poLoTstBuf2 Selection for Enable / Disable
'''
'''@log{2024-02-07, Olia Svyrydova, Initial Version}
'''@log{2025-11-12, Channon Andrew, ported from 8191\, @JIRA{RSIPPTE-108}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeLoTstBuf2En( _
    Optional poLoTstBuf2 As OptLoTstBuf)

    If poLoTstBuf2 Then
        With IfxRegMap.NewRegTxTxloTestStaCf
            .bfEnTestbuf2 = 1
            .bfEnStabuf2Filterbypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
            Call Dopp.Wait(SR_TX_BIAS_EN_FILTERBYPASS_WAIT)
            .bfEnStabuf2Filterbypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    Else
        With IfxRegMap.NewRegTxTxloTestStaCf
            .bfEnTestbuf2 = 0
            .bfEnStabuf2Filterbypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
            Call Dopp.Wait(SR_TX_BIAS_EN_FILTERBYPASS_WAIT)
            .bfEnStabuf2Filterbypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    Call Dopp.Wait(SR_TXLO_LOIN_CONF_EN_WAIT)
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Internal utility Function to retrieve the current module logger
'''@details
'''Intended to be used for displaying error messages and debug event messages as an aid for TP status
'''
'''@return [As ILogger] module logger
'''
'''@log{2025-07-18, Channon Andrew, Initial Version}
'>AUTOCOMMENT:0.4:FUNCTION
Private Function Logger() As ILogger
    Static soLogger As ILogger
    If soLogger Is Nothing Then
        Set soLogger = IfxLog.GetLogger("Program.Tx.Vbt50Txge")
    End If
    Set Logger = soLogger
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief tx noise filter leakage
'''
'''@details
''' Measures in leakage current of 0V9RF and 0V9PA for different filter configurations
'''
'''@param[in] psRelaySetup added to the instance context dictionary in the InitInstance method
'''@param[in] Validating_ Flag to determine validation / test program execution
'''
'''@log{2025-03-05, Leftheriotis Georgios, Initial Version\, @JIRA{CTRXTE-5626}}
'''@log{2025-10-29, Uzun Bilal, Replaced HardReset with EnableDfxGateProt and SoftReset\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeLeakageCurrentMirrorsTXLO( _
    psRelaySetup As String, _
    Optional Validating_ As Boolean = False) As Long
    
    Dim liCounter As Long
    Dim liFilterSet As Long
    
    Call InitInstance(Validating_)
    If Validating_ Then
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Call DeviceSetup
    TheHdw.Utility.Pins("K6_Vdd1v0txToCap").State = tlUtilBitOff
    With TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter
        .Mode = tlDCVIMeterCurrent
        .HardwareAverage = 256
    End With
    
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName & "TxLeakpt1", _
        pbLargeCaptureMode:=True, pbReadFailBitCnt:=False)
    '''-# configure registers for leakage measurements
    mdDcviStrobeWait = IfxHdw.Dcvi.GetStrobeWait("VDD1V0TX_MDCVI")
    Call TxgeTxLeakageSettingInitial
    Call TxgeTXLOFilterConfigLOpt1
    Call Ctrx.DoppStopPattern
    
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName & "TxLeakNoReg", _
        pbLargeCaptureMode:=True, pbReadFailBitCnt:=False)
    Call TxgeTxLeakageSettingInitial
    Call TxgeTXLOFilterConfig3
    Call Ctrx.EnableDfxGateProtection
    Call Ctrx.SoftReset
    Call Ctrx.DoppStopPattern
    
    Call TxgeFilterCurrentReadoutLog2
    
    tfTxgeLeakageCurrentMirrorsTXLO = EndInstance
    
    Exit Function
ErrHandler:
    If AbortTest Then HardTesterReset Else Resume Next
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitial()
    
    ''' -# Enable TX central biasing
    Call TxCentralBias(OptTxCentralBias_En)
    
    ''' -# Enable the TXLO distribution biasing
    With IfxRegMap.NewRegTxTxloBiasCf
        .bfLoEnBias = 1
        .bfEnLodistBias = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    Call Dopp.Wait(SR_TXLO_BIAS_CONF_EN_WAIT)
        
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief filter configurations
'''@details
'''Iterates through different filter configurations and measures current
'''
'''@log{2025-04-02, Leftheriotis Georgios, Initial Version\, @JIRA{CTRXTE-5684}}
'''@log{2025-04-02, Leftheriotis Georgios, for loop changed to 3\, @JIRA{CTRXTE-5900}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTXLOFilterConfigLOpt1()
    
    Dim liFilterSet As Long
    
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting1(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting2(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting3(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting4(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting5(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting6(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting7(liFilterSet)
    Next liFilterSet
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief filter configurations
'''@details
'''Iterates through different filter configurations and measures current
'''
'''@log{2025-04-02, Leftheriotis Georgios, Initial Version\, @JIRA{CTRXTE-5684}}
'''@log{2025-04-02, Leftheriotis Georgios, for loop changed to 3\, @JIRA{CTRXTE-5900}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTXLOFilterConfig3()
    
    Dim liFilterSet As Long
    
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting11(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting12(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting13(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting15(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting17(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting18(liFilterSet)
    Next liFilterSet
    For liFilterSet = 0 To 3
        Call TxgeTxLeakageSetting19(liFilterSet)
    Next liFilterSet
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Function reads the current measurements
'''
'''@details
''' for all different configurations the data is read from the meter and then post-processed
''' logs all current readouts and then the post-processed data
''' FBP0=(TxrfTxLeakageSetting1(1)-TxrfTxLeakageSetting1(0))
''' FBP1=(TxrfTxLeakageSetting1(3)-TxrfTxLeakageSetting1(2))
''' ratio=FBP0/FB1
'''
'''@log{2025-03-10, lef, Initial Version\, @JIRA{CTRXTE-5684}}
'''@log{2025-07-07, Leftheriotis Georgios, adapted number of measurements for leak3\, @JIRA{CTRXTE-5900}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeFilterCurrentReadoutLog2()
    Const N_MEAS As Long = 4
    Const N_LEAK1 As Long = 7 * N_MEAS
    Const N_LEAK_ISD As Long = 4 * N_MEAS
    Const N_LEAK3 As Long = 7 * N_MEAS
    Dim liIdx As Long
    Dim liLeakAllSize As Long
    Dim lpldLeak1 As New PinListData
    Dim lpldLeakISD As New PinListData
    Dim lpldLeak3 As New PinListData
    Dim lwLeak1 As New DSPWave
    Dim lwLeakISD As New DSPWave
    Dim lwLeak3 As New DSPWave
    Dim lwLeakAll As New DSPWave
    Dim lwDeltaEn As New DSPWave
    Dim lwDeltaBp As New DSPWave
    Dim lwDeltaRatio As New DSPWave
    
    For Each Site In TheExec.Sites
        lpldLeak1 = TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Read(tlNoStrobe, N_LEAK1, , tlDCVIMeterReadingFormatArray)
        lpldLeak3 = TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Read(tlNoStrobe, N_LEAK3, , tlDCVIMeterReadingFormatArray)

        lwLeak1 = IfxMath.ConvertTo.DSPWave(lpldLeak1.Pins("VDD1V0TX_MDCVI"))
        lwLeak3 = IfxMath.ConvertTo.DSPWave(lpldLeak3.Pins("VDD1V0TX_MDCVI"))
    Next Site
    
    If TheExec.TesterMode = testModeOffline Then
        For Each Site In TheExec.Sites
            lwLeak1.FileImport (".\Program\Tx\Waves\TxrfCurLeakMirror1.txt")
            lwLeak3.FileImport (".\Program\Tx\Waves\TxrfCurLeakMirror3.txt")
        Next Site
    End If
    
    For Each Site In TheExec.Sites
        lwLeakAll = lwLeak1.Concatenate(lwLeak3)
    Next Site
    
    liLeakAllSize = N_LEAK1 + N_LEAK3
    
    For liIdx = 0 To liLeakAllSize - 1
        Call IfxEnv.Datalog(lwLeakAll.Element(liIdx))
    Next liIdx
    
    For Each Site In TheExec.Sites
        Call CalcLeakMirrors(lwLeakAll, liLeakAllSize, lwDeltaEn, lwDeltaBp, lwDeltaRatio)
    Next Site
    
    For liIdx = 0 To (liLeakAllSize / N_MEAS) - 1
        Call IfxEnv.Datalog(lwDeltaEn.Element(liIdx))
        Call IfxEnv.Datalog(lwDeltaBp.Element(liIdx))
        Call IfxEnv.Datalog(lwDeltaRatio.Element(liIdx))
    Next liIdx
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: LOOUT1, cFBP:0, cEN:0   REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, Leftheriotis Georgios, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting1(ByRef piTxFilterSet As Long)
    
    Dim loTxTxloLooutCf As RegTxTxloLooutCf
    Set loTxTxloLooutCf = IfxRegMap.NewRegTxTxloLooutCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout1FilterbypassS1(0).withBfEnLoout1FilterbypassS2(0) _
            .withBfEnLoout1Buf(0).withBfEnLoout1DataReady(1))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout1FilterbypassS1(0).withBfEnLoout1FilterbypassS2(0) _
            .withBfEnLoout1Buf(1).withBfEnLoout1DataReady(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout1FilterbypassS1(1).withBfEnLoout1FilterbypassS2(1) _
            .withBfEnLoout1Buf(0).withBfEnLoout1DataReady(1))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout1FilterbypassS1(1).withBfEnLoout1FilterbypassS2(1) _
            .withBfEnLoout1Buf(1).withBfEnLoout1DataReady(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
        .withBfEnLoout1FilterbypassS1(0).withBfEnLoout1FilterbypassS2(0) _
        .withBfEnLoout1Buf(0).withBfEnLoout1DataReady(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: LOOUT2, cFBP:0, cEN:0 REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting2(ByRef piTxFilterSet As Long)
    
    Dim loTxTxloLooutCf As RegTxTxloLooutCf
    Set loTxTxloLooutCf = IfxRegMap.NewRegTxTxloLooutCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout2FilterbypassS1(0).withBfEnLoout2FilterbypassS2(0) _
            .withBfEnLoout2Buf(0).withBfEnLoout2DataReady(1))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout2FilterbypassS1(0).withBfEnLoout2FilterbypassS2(0) _
            .withBfEnLoout2Buf(1).withBfEnLoout2DataReady(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout2FilterbypassS1(1).withBfEnLoout2FilterbypassS2(1) _
            .withBfEnLoout2Buf(0).withBfEnLoout2DataReady(1))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
            .withBfEnLoout2FilterbypassS1(1).withBfEnLoout2FilterbypassS2(1) _
            .withBfEnLoout2Buf(1).withBfEnLoout2DataReady(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(loTxTxloLooutCf _
        .withBfEnLoout2FilterbypassS1(0).withBfEnLoout2FilterbypassS2(0) _
        .withBfEnLoout2Buf(0).withBfEnLoout2DataReady(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: SSD_TXLO, cFBP:0, cEN:0 REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting3(ByRef piTxFilterSet As Long)
    
    Dim loRegTxTxloLodistCf As RegTxTxloLodistCf
    Set loRegTxTxloLodistCf = IfxRegMap.NewRegTxTxloLodistCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnSsdTxlo(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnSsdTxlo(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnSsdTxlo(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnSsdTxlo(1))
    End If
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister( _
        loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnSsdTxlo(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: ISD_TX14, cFBP:0, cEN:0   REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting4(ByRef piTxFilterSet As Long)
    
    Dim loRegTxTxloLodistCf As RegTxTxloLodistCf
    Set loRegTxTxloLodistCf = IfxRegMap.NewRegTxTxloLodistCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxLeft(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxLeft(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnAmpTxLeft(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnAmpTxLeft(1))
    End If
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister( _
        loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxLeft(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: ISD_TX48, cFBP:0, cEN:0   REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting5(ByRef piTxFilterSet As Long)
    
    Dim loRegTxTxloLodistCf As RegTxTxloLodistCf
    Set loRegTxTxloLodistCf = IfxRegMap.NewRegTxTxloLodistCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxRight(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxRight(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnAmpTxRight(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnAmpTxRight(1))
    End If
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister( _
        loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxRight(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: LOSPLIT_ISD_TXMON, cFBP:0, cEN:0  REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting6(ByRef piTxFilterSet As Long)
    
    Dim loRegTxTxloLodistCf As RegTxTxloLodistCf
    Set loRegTxTxloLodistCf = IfxRegMap.NewRegTxTxloLodistCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxRx(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxRx(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnAmpTxRx(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister( _
            loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnAmpTxRx(1))
    End If
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister( _
        loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnAmpTxRx(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: SSD_RXLO, cFBP:0, cEN:0 REGULATED
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting7(ByRef piTxFilterSet As Long)
    
    Dim loRegTxTxloLodistCf As RegTxTxloLodistCf
    Set loRegTxTxloLodistCf = IfxRegMap.NewRegTxTxloLodistCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnSsdRxlo(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnSsdRxlo(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnSsdRxlo(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(loRegTxTxloLodistCf.withBfEnFilterbypass(1).withBfEnSsdRxlo(1))
    End If
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(loRegTxTxloLodistCf.withBfEnFilterbypass(0).withBfEnSsdRxlo(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: FQM2
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-07-07, Leftheriotis Georgios, removed case 4\, @JIRA{CTRXTE-5900}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting11(ByRef piTxFilterSet As Long)
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm2Cf _
            .withBfEnFilterbypass(0).withBfEnBias(0).withBfEnMirror(0).withBfEnBiasCore(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm2Cf _
            .withBfEnFilterbypass(0).withBfEnBias(1).withBfEnMirror(1).withBfEnBiasCore(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm2Cf _
            .withBfEnFilterbypass(1).withBfEnBias(0).withBfEnMirror(0).withBfEnBiasCore(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm2Cf _
            .withBfEnFilterbypass(1).withBfEnBias(1).withBfEnMirror(1).withBfEnBiasCore(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm2Cf _
        .withBfEnFilterbypass(0).withBfEnBias(0).withBfEnMirror(0).withBfEnBiasCore(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: FQM3
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting12(ByRef piTxFilterSet As Long)
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm3Cf.withBfEnFilterbypass(0). _
                                                                withBfEnBias(0). _
                                                                withBfEnMirror(0). _
                                                                withBfEnBiasBuf2(0). _
                                                                withBfEnBiasBuf3(0). _
                                                                withBfEnBiasCore(0))

    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm3Cf.withBfEnFilterbypass(0). _
                                                                withBfEnBias(1). _
                                                                withBfEnMirror(1). _
                                                                withBfEnBiasBuf2(1). _
                                                                withBfEnBiasBuf3(1). _
                                                                withBfEnBiasCore(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm3Cf.withBfEnFilterbypass(1). _
                                                                withBfEnBias(0). _
                                                                withBfEnMirror(0). _
                                                                withBfEnBiasBuf2(0). _
                                                                withBfEnBiasBuf3(0). _
                                                                withBfEnBiasCore(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm3Cf.withBfEnFilterbypass(1). _
                                                                withBfEnBias(1). _
                                                                withBfEnMirror(1). _
                                                                withBfEnBiasBuf2(1). _
                                                                withBfEnBiasBuf3(1). _
                                                                withBfEnBiasCore(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloFqm3Cf.withBfEnFilterbypass(0). _
                                                                withBfEnBias(0). _
                                                                withBfEnMirror(0). _
                                                                withBfEnBiasBuf2(0). _
                                                                withBfEnBiasBuf3(0). _
                                                                withBfEnBiasCore(0))
    
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: CASSPLIT1
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting13(ByRef piTxFilterSet As Long)
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf.withBfEnFilterbypass(0).withBfEnBuf(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf.withBfEnFilterbypass(0).withBfEnBuf(1) _
            .withBfVgbiasEn(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf.withBfEnFilterbypass(1).withBfEnBuf(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf.withBfEnFilterbypass(1).withBfEnBuf(1) _
            .withBfVgbiasEn(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf.withBfEnFilterbypass(0).withBfEnBuf(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: CASSPLIT2
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting14(ByRef piTxFilterSet As Long)
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf _
            .withBfEnFilterbypass(0).withBfEnBuf(0).withBfSelBuf(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf _
            .withBfEnFilterbypass(0).withBfEnBuf(1).withBfSelBuf(1).withBfVgbiasEn(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf _
            .withBfEnFilterbypass(1).withBfEnBuf(0).withBfSelBuf(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf _
            .withBfEnFilterbypass(1).withBfEnBuf(1).withBfSelBuf(1).withBfVgbiasEn(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloCassplitCf _
        .withBfEnFilterbypass(0).withBfEnBuf(0).withBfSelBuf(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: CASCOMB1
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting15(ByRef piTxFilterSet As Long)
    
    Dim loTxTxloCascombCf As RegTxTxloCascombCf
    Set loTxTxloCascombCf = IfxRegMap.NewRegTxTxloCascombCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(loTxTxloCascombCf _
            .withBfEnFilterbypass(0).withBfEnBuf(0).withBfEnComb(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(loTxTxloCascombCf _
            .withBfEnFilterbypass(0).withBfEnBuf(1).withBfEnComb(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(loTxTxloCascombCf _
            .withBfEnFilterbypass(1).withBfEnBuf(0).withBfEnComb(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(loTxTxloCascombCf _
            .withBfEnFilterbypass(1).withBfEnBuf(1).withBfEnComb(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(loTxTxloCascombCf _
        .withBfEnFilterbypass(0).withBfEnBuf(0).withBfEnComb(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: LOIN1
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting17(ByRef piTxFilterSet As Long)
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin1Filterbypass(0).withBfEnLoin1Buf(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin1Filterbypass(0).withBfEnLoin1Buf(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin1Filterbypass(1).withBfEnLoin1Buf(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin1Filterbypass(1).withBfEnLoin1Buf(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin1Filterbypass(0).withBfEnLoin1Buf(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: LOIN2
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting18(ByRef piTxFilterSet As Long)
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin2Filterbypass(0).withBfEnLoin2Buf(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin2Filterbypass(0).withBfEnLoin2Buf(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin2Filterbypass(1).withBfEnLoin2Buf(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin2Filterbypass(1).withBfEnLoin2Buf(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxloLoinCf.withBfEnLoin2Filterbypass(0).withBfEnLoin2Buf(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to measure Block: STABUF1
'''
'''@param[in] piTxFilterSet The filter number
'''
'''@log{2025-03-06, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSetting19(ByRef piTxFilterSet As Long)
    
    Dim loTxTxloTestStaCf As RegTxTxloTestStaCf
    Set loTxTxloTestStaCf = IfxRegMap.NewRegTxTxloTestStaCf
    
    If piTxFilterSet = 0 Then
        Call Ctrx.WriteFixedRegister(loTxTxloTestStaCf.withBfEnStabuf1Filterbypass(0).withBfEnStabuf1(0))
    ElseIf piTxFilterSet = 1 Then
        Call Ctrx.WriteFixedRegister(loTxTxloTestStaCf.withBfEnStabuf1Filterbypass(0).withBfEnStabuf1(1))
    ElseIf piTxFilterSet = 2 Then
        Call Ctrx.WriteFixedRegister(loTxTxloTestStaCf.withBfEnStabuf1Filterbypass(1).withBfEnStabuf1(0))
    ElseIf piTxFilterSet = 3 Then
        Call Ctrx.WriteFixedRegister(loTxTxloTestStaCf.withBfEnStabuf1Filterbypass(1).withBfEnStabuf1(1))
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    Call Ctrx.WriteFixedRegister(loTxTxloTestStaCf.withBfEnStabuf1Filterbypass(0).withBfEnStabuf1(0))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
''' @brief Calculate the difference values for CurWFiltEn and CurWFiltBp then calculate the ratios.
''' @details The Corrected Value = VDD + Offset. And the Offset = VDD - PreCalRead.
''' So for or calculation this gives us:
'''  Corrected value = VDD + VDD - PreCalRead
'''
''' @param[in] pwRawValues The measured current leakage values at the supply DCVIs
''' @param[in] pdRawSize The size of the RawValues array which must be divisible by 4.
''' @param[out] pwDeltaEn The difference array for WFiltEn current measurements
''' @param[out] pwDeltaBp The difference array for WFiltBp current measurements
''' @param[out] pwDeltaRatio The calculated ratio of the En value divided by the Bp value
'''
''' @log{2025-04-04, Channon Andrew, @JIRA{CTRXTE-5684}: initial version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function CalcLeakMirrors( _
    ByVal pwRawValues As DSPWave, _
    ByVal piRawSize As Long, _
    ByRef pwDeltaEn As DSPWave, _
    ByRef pwDeltaBp As DSPWave, _
    ByRef pwDeltaRatio As DSPWave) As Long
    
    Const N_MEAS As Long = 4
    Const NON_ZERO As Double = 5
    Dim liResultSize As Long
    Dim lwCurWFiltEnA As New DSPWave
    Dim lwCurWFiltEnB As New DSPWave
    Dim lwCurWFiltBpA As New DSPWave
    Dim lwCurWFiltBpB As New DSPWave
    Dim lwZeroValues As New DSPWave
    Dim lwNonZero As New DSPWave
    
    ''' we expect the DspWave array size to match the RawSize parameter and this must be divisible by 4
    ''' - the RawValues contain multiple sets of 4 current measurements each
    If pwRawValues.SampleSize <> piRawSize Then
        'Measurement error -> likely a failure in the measured values which fails the test
        CalcLeakMirrors = TL_ERROR
        Exit Function
    End If
    
    ''' separate out each of the 4 measurement types
    ''' - The first 2 relate to WFiltEn current measurements
    ''' - The second 2 relate to WFiltBp current measurements
    lwCurWFiltEnA = pwRawValues.Select(0, N_MEAS)
    lwCurWFiltEnB = pwRawValues.Select(1, N_MEAS)
    lwCurWFiltBpA = pwRawValues.Select(2, N_MEAS)
    lwCurWFiltBpB = pwRawValues.Select(3, N_MEAS)
    
    ''' Calculate the En and Bp current deltas as (B-A) and return these two results arrays
    liResultSize = piRawSize / N_MEAS
    Call pwDeltaEn.CreateConstant(0, liResultSize, DspDouble)
    Call pwDeltaBp.CreateConstant(0, liResultSize, DspDouble)
    pwDeltaEn = lwCurWFiltEnB.Subtract(lwCurWFiltEnA)
    pwDeltaBp = lwCurWFiltBpB.Subtract(lwCurWFiltBpA)
    
    ''' Calculate the ratio CurWFiltEn divided by CurWFiltBp and return this results array also
    ''' - includes check and replacement for elements with value 0 to avoid math error
    lwZeroValues = pwDeltaBp.FindIndices(EqualTo, 0)
    Call lwNonZero.CreateConstant(NON_ZERO, liResultSize, DspDouble)
    Call pwDeltaBp.ReplaceElements(lwZeroValues, lwNonZero)
    
    Call pwDeltaRatio.CreateConstant(0, liResultSize, DspDouble)
    pwDeltaRatio = pwDeltaEn.Divide(pwDeltaBp)
    
    CalcLeakMirrors = TL_SUCCESS
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-03-26, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh1()
    
    Dim loTxTxrf1DigCf As RegTxTxrf1DigCf
    Set loTxTxrf1DigCf = IfxRegMap.NewRegTxTxrf1DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf1PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf1BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-03-26, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh2()
    
    Dim loTxTxrf2DigCf As RegTxTxrf2DigCf
    Set loTxTxrf2DigCf = IfxRegMap.NewRegTxTxrf2DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf2PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf2BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-03-26, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh3()
    
    Dim loTxTxrf3DigCf As RegTxTxrf3DigCf
    Set loTxTxrf3DigCf = IfxRegMap.NewRegTxTxrf3DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf3PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf3BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-03-26, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh4()
    
    Dim loTxTxrf4DigCf As RegTxTxrf4DigCf
    Set loTxTxrf4DigCf = IfxRegMap.NewRegTxTxrf4DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf4PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf4BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-11-26, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh5()
    
    Dim loTxTxrf5DigCf As RegTxTxrf5DigCf
    Set loTxTxrf5DigCf = IfxRegMap.NewRegTxTxrf5DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf5PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf5BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-11-26, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh6()
    
    Dim loTxTxrf6DigCf As RegTxTxrf6DigCf
    Set loTxTxrf6DigCf = IfxRegMap.NewRegTxTxrf6DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf6PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf6BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-11-26, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh7()
    
    Dim loTxTxrf7DigCf As RegTxTxrf7DigCf
    Set loTxTxrf7DigCf = IfxRegMap.NewRegTxTxrf7DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf7PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf7BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to Perform DC current measurement - TXPA
'''
'''@log{2025-11-26, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPACh8()
    
    Dim loTxTxrf8DigCf As RegTxTxrf8DigCf
    Set loTxTxrf8DigCf = IfxRegMap.NewRegTxTxrf8DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf.withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPaDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPaDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf8PaCf.withBfEnIdac(0))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf8BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPaDigInit(1).withBfPaDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPaDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PA Stage: 1,TX leakage measurement   TxTxrf1PaFilterRwh
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPAStage1( _
    ByRef piTxFilterSet As Long, _
    ByRef piTxChan As Long)
    
    With IfxRegMap.NewRegTxTxrf1PaCf
        .bfEnIdac = 1
        .bfEnClassb = 1
        .bfEnStage1 = 1
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa0FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 1
            .bfEnStage1 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa0Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa0FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 1
            .bfEnStage1 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa0Idac = 200
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa0FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 1
            .bfEnStage1 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa0Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa0FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 1
            .bfEnStage1 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa0Idac = 200
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
    Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    
    If piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa0FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 0
            .bfEnStage1 = 0
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa0Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PA Stage: 2,TX leakage measurement   TxTxrf1PaFilterRwh
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPAStage2( _
    ByRef piTxFilterSet As Long, _
    ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa1FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 2
            .bfEnStage2 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa1Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa1FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 2
            .bfEnStage2 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa1Idac = 160
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa1FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 2
            .bfEnStage2 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa1Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa1FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 2
            .bfEnStage2 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa1Idac = 160
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
    Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    
    If piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa1FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 0
            .bfEnStage2 = 0
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa1Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PA Stage: 3,TX leakage measurement   TxTxrf1PaFilterRwh
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPAStage3( _
    ByRef piTxFilterSet As Long, _
    ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa2FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 4
            .bfEnStage3 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa2Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa2FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 4
            .bfEnStage3 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa2Idac = 220
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa2FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 4
            .bfEnStage3 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa2Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa2FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 4
            .bfEnStage3 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa2Idac = 220
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
    Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    
    If piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa2FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 0
            .bfEnStage3 = 0
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa2Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PA Stage: 4,TX leakage measurement   TxTxrf1PaFilterRwh
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPAStage4( _
    ByRef piTxFilterSet As Long, _
    ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa3FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 8
            .bfEnStage4 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa3Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa3FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 8
            .bfEnStage4 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa3Idac = 255
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa3FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 8
            .bfEnStage4 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa3Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa3FilterBypass = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 8
            .bfEnStage4 = 1
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa3Idac = 255
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    Call Dopp.Wait(WAIT_TOTSYS_SETTLE)
    Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    
    If piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaFilterRwh", VbCallType.VbMethod)
            .bfPa3FilterBypass = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaCf", VbCallType.VbMethod)
            .bfEnIdac = 0
            .bfEnStage4 = 0
            .bfEnClassb = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PaIdacRwh", VbCallType.VbMethod)
            .bfPa3Idac = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Function reads the current measurements
'''
'''@details
''' for all different configurations the data is read from the meter and then post-processed
''' logs all current readouts and then the post-processed data
''' FBP0=(TxrfTxLeakageSetting1(1)-TxrfTxLeakageSetting1(0))
''' FBP1=(TxrfTxLeakageSetting1(3)-TxrfTxLeakageSetting1(2))
''' ratio=FBP0/FB1
'''
'''@param[in] pbMerged Flag for merged measurement
'''
'''@log{2025-04-08, lef, Initial Version\, @JIRA{CTRXTE-5731}}
'''@log{2025-11-25, lef, Added merged boolean variable and adapted for merged domains\, @JIRA{CTRXTE-6365}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeFilterCurrentReadoutPALog2(ByRef pbMerged As Boolean)
    Const N_MEAS As Long = 2
    Const N_TXCH As Long = 8
    Dim liNLeakTxPA As Long
    Dim liIdx As Long
    Dim liLeakAllSize As Long
    Dim lpldLeakTXPA(0 To 7) As New PinListData
    Dim lwLeakTXPA(0 To 7) As New DSPWave
    Dim lwLeakTXPATot As New DSPWave
    Dim lwDeltaEn1 As New DSPWave
    Dim lwDeltaBp1 As New DSPWave
    Dim lwDeltaRatio1 As New DSPWave
    Dim lwDeltaEn2 As New DSPWave
    Dim lwDeltaBp2 As New DSPWave
    Dim lwDeltaRatio2 As New DSPWave
    Dim lwDeltaEn3 As New DSPWave
    Dim lwDeltaBp3 As New DSPWave
    Dim lwDeltaRatio3 As New DSPWave
    Dim lwDeltaEn4 As New DSPWave
    Dim lwDeltaBp4 As New DSPWave
    Dim lwDeltaRatio4 As New DSPWave
    Dim lwDeltaEn5 As New DSPWave
    Dim lwDeltaBp5 As New DSPWave
    Dim lwDeltaRatio5 As New DSPWave
    Dim lwDeltaEn6 As New DSPWave
    Dim lwDeltaBp6 As New DSPWave
    Dim lwDeltaRatio6 As New DSPWave
    Dim lwDeltaEn7 As New DSPWave
    Dim lwDeltaBp7 As New DSPWave
    Dim lwDeltaRatio7 As New DSPWave
    Dim lwDeltaEn8 As New DSPWave
    Dim lwDeltaBp8 As New DSPWave
    Dim lwDeltaRatio8 As New DSPWave
    Dim lwDeltaEn As New DSPWave
    Dim lwDeltaBp As New DSPWave
    Dim lwDeltaRatio As New DSPWave
    Dim liCtr As Long
    If pbMerged Then
        liNLeakTxPA = 4 * N_MEAS
    Else
        liNLeakTxPA = 4 * (N_MEAS * 2)
    End If
    
    For Each Site In TheExec.Sites
        For liCtr = 0 To 7
            lpldLeakTXPA(liCtr) = TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Read(tlNoStrobe, liNLeakTxPA, , _
                tlDCVIMeterReadingFormatArray)
            lwLeakTXPA(liCtr) = IfxMath.ConvertTo.DSPWave(lpldLeakTXPA(liCtr).Pins("VDD1V0TX_MDCVI"))
        Next liCtr
    Next Site
    
    If TheExec.TesterMode = testModeOffline Then
        If pbMerged Then
            'Merged measurement is not yet available, to be enabled in future if needed
        Else
            For Each Site In TheExec.Sites
                For liCtr = 0 To 7
                    lwLeakTXPA(liCtr).FileImport _
                        (".\Program\Tx\Waves\TxrfCurLeakMirrorTXPA" & CStr(1 + (liCtr Mod 4)) & ".txt")
                Next liCtr
            Next Site
        End If
    End If
    liLeakAllSize = liNLeakTxPA
    For Each Site In TheExec.Sites
        lwLeakTXPATot = lwLeakTXPA(0).Concatenate(lwLeakTXPA(1)).Concatenate(lwLeakTXPA(2)).Concatenate(lwLeakTXPA(3)) _
            .Concatenate(lwLeakTXPA(4)).Concatenate(lwLeakTXPA(5)).Concatenate(lwLeakTXPA(6)).Concatenate(lwLeakTXPA(7))
    Next Site
    For liIdx = 0 To liLeakAllSize * N_TXCH - 1
        Call IfxEnv.Datalog(lwLeakTXPATot.Element(liIdx))
    Next liIdx
    
    For Each Site In TheExec.Sites
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(0), liLeakAllSize, lwDeltaEn1, lwDeltaBp1, lwDeltaRatio1)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(1), liLeakAllSize, lwDeltaEn2, lwDeltaBp2, lwDeltaRatio2)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(2), liLeakAllSize, lwDeltaEn3, lwDeltaBp3, lwDeltaRatio3)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(3), liLeakAllSize, lwDeltaEn4, lwDeltaBp4, lwDeltaRatio4)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(4), liLeakAllSize, lwDeltaEn5, lwDeltaBp5, lwDeltaRatio5)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(5), liLeakAllSize, lwDeltaEn6, lwDeltaBp6, lwDeltaRatio6)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(6), liLeakAllSize, lwDeltaEn7, lwDeltaBp7, lwDeltaRatio7)
        Call CalcLeakMirrorsTXPA(lwLeakTXPA(7), liLeakAllSize, lwDeltaEn8, lwDeltaBp8, lwDeltaRatio8)
    Next Site
    
    For Each Site In TheExec.Sites
        lwDeltaEn = lwDeltaEn1.Concatenate(lwDeltaEn2).Concatenate(lwDeltaEn3).Concatenate(lwDeltaEn4) _
            .Concatenate(lwDeltaEn5).Concatenate(lwDeltaEn6).Concatenate(lwDeltaEn7).Concatenate(lwDeltaEn8)
        lwDeltaBp = lwDeltaBp1.Concatenate(lwDeltaBp2).Concatenate(lwDeltaBp3).Concatenate(lwDeltaBp4) _
            .Concatenate(lwDeltaBp5).Concatenate(lwDeltaBp6).Concatenate(lwDeltaBp7).Concatenate(lwDeltaBp8)
        lwDeltaRatio = lwDeltaRatio1.Concatenate(lwDeltaRatio2).Concatenate(lwDeltaRatio3).Concatenate(lwDeltaRatio4) _
            .Concatenate(lwDeltaRatio5).Concatenate(lwDeltaRatio6).Concatenate(lwDeltaRatio7).Concatenate(lwDeltaRatio8)
    Next Site
    For liIdx = 0 To (liNLeakTxPA * 2) - 1
        Call IfxEnv.Datalog(lwDeltaEn.Element(liIdx))
        Call IfxEnv.Datalog(lwDeltaBp.Element(liIdx))
        Call IfxEnv.Datalog(lwDeltaRatio.Element(liIdx))
    Next liIdx
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief tx noise filter leakage
'''
'''@details
''' Measures in leakage current of 0V9PA for different filter configurations
'''
'''@param[in] psRelaySetup added to the instance context dictionary in the InitInstance method
'''@param[in] pbMerged boolean to decide measuring TXPA stages merged or not
'''@param[in] Validating_ Flag to determine validation / test program execution
'''
'''@log{2025-03-31, Leftheriotis Georgios, Initial Version\, @JIRA{CTRXTE-5626}}
'''@log{2025-10-29, Uzun Bilal, Replaced HardReset with EnableDfxGateProt and SoftReset\, @JIRA{CTRXTE-6008}}
'''@log{2025-11-25, lef, Added merged boolean variable and adapted for merged domains\, @JIRA{CTRXTE-6365}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeLeakageCurrentMirrorsTXPA( _
    psRelaySetup As String, _
    Optional pbMerged As Boolean = False, _
    Optional Validating_ As Boolean = False) As Long
    
    Dim liCounter As Long
    Dim liChannelSet As Long
    Dim liFilterSet As Long
    
    Call InitInstance(Validating_, pbLogMeanOffline:=False)
    If Validating_ Then
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Call DeviceSetup
    TheHdw.Utility.Pins("K6_Vdd1v0txToCap").State = tlUtilBitOff
    
    With TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter
        .Mode = tlDCVIMeterCurrent
        .HardwareAverage = 256
    End With
    mdDcviStrobeWait = IfxHdw.Dcvi.GetStrobeWait("VDD1V0TX_MDCVI")
    
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName & "LeakTXPA" & pbMerged, _
        pbLargeCaptureMode:=True, pbReadFailBitCnt:=False)
    For liChannelSet = 1 To 8
        '''-# Disable any DFXgate protection
        Call Ctrx.DisableDfxGateProtection
        If liChannelSet = 1 Then
            Call TxgeTxLeakageSettingInitialTXPACh1
        ElseIf liChannelSet = 2 Then
            Call TxgeTxLeakageSettingInitialTXPACh2
        ElseIf liChannelSet = 3 Then
            Call TxgeTxLeakageSettingInitialTXPACh3
        ElseIf liChannelSet = 4 Then
            Call TxgeTxLeakageSettingInitialTXPACh4
        ElseIf liChannelSet = 5 Then
            Call TxgeTxLeakageSettingInitialTXPACh5
        ElseIf liChannelSet = 6 Then
            Call TxgeTxLeakageSettingInitialTXPACh6
        ElseIf liChannelSet = 7 Then
            Call TxgeTxLeakageSettingInitialTXPACh7
        ElseIf liChannelSet = 8 Then
            Call TxgeTxLeakageSettingInitialTXPACh8
        End If
        
        For liFilterSet = 0 To 3
            If pbMerged Then
                'Merged measurement is not yet available, to be enabled in future if needed
            Else
                Call TxgeTxLeakageSettingPAStage1(liFilterSet, liChannelSet)
                Call TxgeTxLeakageSettingPAStage2(liFilterSet, liChannelSet)
                Call TxgeTxLeakageSettingPAStage3(liFilterSet, liChannelSet)
                Call TxgeTxLeakageSettingPAStage4(liFilterSet, liChannelSet)
            End If
        Next liFilterSet
        Call Ctrx.EnableDfxGateProtection
        Call Ctrx.SoftReset
    Next liChannelSet
    Call Ctrx.DoppStopPattern
    Call TxgeFilterCurrentReadoutPALog2(pbMerged)
    
    tfTxgeLeakageCurrentMirrorsTXPA = EndInstance()
    Exit Function
ErrHandler:
    If AbortTest Then HardTesterReset Else Resume Next
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief tx noise filter leakage
'''
'''@details
''' Measures in leakage current of 0V9PA for different filter configurations
'''
'''@param[in] psRelaySetup added to the instance context dictionary in the InitInstance method
'''@param[in] Validating_ Flag to determine validation / test program execution
'''
'''@log{2025-03-31, Leftheriotis Georgios, Initial Version\, @JIRA{CTRXTE-5626}}
'''@log{2025-10-29, Uzun Bilal, Replaced HardReset with EnableDfxGateProt and SoftReset\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function tfTxgeLeakageCurrentMirrorsTXPS( _
    psRelaySetup As String, _
    Optional Validating_ As Boolean = False) As Long
    
    Dim liCounter As Long
    Dim liChannelSet As Long
    Dim liFilterSet As Long
    
    Call InitInstance(Validating_, pbLogMeanOffline:=False)
    If Validating_ Then
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    Call DeviceSetup
    TheHdw.Utility.Pins("K6_Vdd1v0txToCap").State = tlUtilBitOff
    
    With TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter
        .Mode = tlDCVIMeterCurrent
        .HardwareAverage = 256
    End With
    mdDcviStrobeWait = IfxHdw.Dcvi.GetStrobeWait("VDD1V0TX_MDCVI")
    
    Call Ctrx.DoppStartPattern(TX_DOPP_PATH & GetPatternName & "LeakTXPS", _
        pbLargeCaptureMode:=True, pbReadFailBitCnt:=False)
    For liChannelSet = 1 To 8
        '''-# Disable any DFXgate protection
        Call Ctrx.DisableDfxGateProtection
        If liChannelSet = 1 Then
            Call TxgeTxLeakageSettingInitialTXPSTX1
        ElseIf liChannelSet = 2 Then
            Call TxgeTxLeakageSettingInitialTXPSTX2
        ElseIf liChannelSet = 3 Then
            Call TxgeTxLeakageSettingInitialTXPSTX3
        ElseIf liChannelSet = 4 Then
            Call TxgeTxLeakageSettingInitialTXPSTX4
        ElseIf liChannelSet = 5 Then
            Call TxgeTxLeakageSettingInitialTXPSTX5
        ElseIf liChannelSet = 6 Then
            Call TxgeTxLeakageSettingInitialTXPSTX6
        ElseIf liChannelSet = 7 Then
            Call TxgeTxLeakageSettingInitialTXPSTX7
        ElseIf liChannelSet = 8 Then
            Call TxgeTxLeakageSettingInitialTXPSTX8
        End If
        
        For liFilterSet = 0 To 3
            Call TxgeTxLeakageSettingPS_IB1Q_TX(liFilterSet, liChannelSet)
            Call TxgeTxLeakageSettingPS_IB1I_TX(liFilterSet, liChannelSet)
            Call TxgeTxLeakageSettingPS_IB2_TX(liFilterSet, liChannelSet)
            Call TxgeTxLeakageSettingPS_IDACQ_TX(liFilterSet, liChannelSet)
            Call TxgeTxLeakageSettingPS_IDACI_TX(liFilterSet, liChannelSet)
        Next liFilterSet
        Call Ctrx.EnableDfxGateProtection
        Call Ctrx.SoftReset
    Next liChannelSet
    Call Ctrx.DoppStopPattern
    
    Call TxgeFilterCurrentReadoutPSLog2
    
    tfTxgeLeakageCurrentMirrorsTXPS = EndInstance
    
    Exit Function
ErrHandler:
    If AbortTest Then HardTesterReset Else Resume Next
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX1()
    
    Dim loTxTxrf1DigCf As RegTxTxrf1DigCf
    Set loTxTxrf1DigCf = IfxRegMap.NewRegTxTxrf1DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf1BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf1DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX2()
    
    Dim loTxTxrf2DigCf As RegTxTxrf2DigCf
    Set loTxTxrf2DigCf = IfxRegMap.NewRegTxTxrf2DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf2BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf2DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX3()
    
    Dim loTxTxrf3DigCf As RegTxTxrf3DigCf
    Set loTxTxrf3DigCf = IfxRegMap.NewRegTxTxrf3DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf3BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf3DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Volatile to fixed reg write conversion for TTR\, @JIRA{CTRXTE-6008}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX4()
    
    Dim loTxTxrf4DigCf As RegTxTxrf4DigCf
    Set loTxTxrf4DigCf = IfxRegMap.NewRegTxTxrf4DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf4BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf4DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Extended functions to 8 TX channels\, @JIRA{RSIPPTE-542}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX5()
    
    Dim loTxTxrf5DigCf As RegTxTxrf5DigCf
    Set loTxTxrf5DigCf = IfxRegMap.NewRegTxTxrf5DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf5BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf5DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Extended functions to 8 TX channels\, @JIRA{RSIPPTE-542}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX6()
    
    Dim loTxTxrf6DigCf As RegTxTxrf6DigCf
    Set loTxTxrf6DigCf = IfxRegMap.NewRegTxTxrf6DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf6BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf6DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Extended functions to 8 TX channels\, @JIRA{RSIPPTE-542}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX7()
    
    Dim loTxTxrf7DigCf As RegTxTxrf7DigCf
    Set loTxTxrf7DigCf = IfxRegMap.NewRegTxTxrf7DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf7BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf7DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable central LO distribution bias
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2025-10-29, Uzun Bilal, Extended functions to 8 TX channels\, @JIRA{RSIPPTE-542}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingInitialTXPSTX8()
    
    Dim loTxTxrf8DigCf As RegTxTxrf8DigCf
    Set loTxTxrf8DigCf = IfxRegMap.NewRegTxTxrf8DigCf
    
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxbiasBiasCf _
        .withBfEnVref(1).withBfEnVrefFilt(1).withBfEnV2i(1))
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPsDigInit(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPsDigInit(0).withBfErrEn(1))
    Call Ctrx.WriteFixedRegister(IfxRegMap.NewRegTxTxrf8BiasCf.withBfEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPsDigInit(1).withBfPsDigEn(1).withBfErrEn(1))
    Call Dopp.Wait(500 * us)
    Call Ctrx.WriteFixedRegister(loTxTxrf8DigCf.withBfPsDigInit(0).withBfErrEn(1))
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PS_IB1Q,TX1 leakage measurement
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPS_IB1Q_TX(ByRef piTxFilterSet As Long, ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnIb1q = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnIb1q = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnIb1q = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnIb1q = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_ISD_PA_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
        .bfEnFilterbypass = 0
        .bfEnIb1q = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PS_IB1I, TXx leakage measurement
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPS_IB1I_TX(ByRef piTxFilterSet As Long, ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnIb1i = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnIb1i = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnIb1i = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnIb1i = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_ISD_PA_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    
    With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
        .bfEnFilterbypass = 0
        .bfEnIb1q = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
    
    With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
        .bfEnFilterbypass = 0
        .bfEnIb1i = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PS_IB2,TX1 leakage measurement
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPS_IB2_TX(ByRef piTxFilterSet As Long, ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnIb2 = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnIb2 = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnIb2 = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnIb2 = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_ISD_PA_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
        .bfEnFilterbypass = 0
        .bfEnIb2 = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PS_IDACQ,TX1 leakage measurement
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPS_IDACQ_TX(ByRef piTxFilterSet As Long, ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnBiasIdacq = 0
            .bfEnVmq = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnBiasIdacq = 1
            .bfEnVmq = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnBiasIdacq = 0
            .bfEnVmq = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnBiasIdacq = 1
            .bfEnVmq = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_ISD_PA_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
        .bfEnFilterbypass = 0
        .bfEnBiasIdacq = 0
        .bfEnVmq = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief settings for tx leakage measurement
'''@details
'''It sets the needed registers to enable PS_IDACI,TX1 leakage measurement
'''
'''@param[in] piTxFilterSet The filter number
'''@param[in] piTxChan The channel number
'''
'''@log{2025-03-24, lef, Initial Version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeTxLeakageSettingPS_IDACI_TX(ByRef piTxFilterSet As Long, ByRef piTxChan As Long)
    
    If piTxFilterSet = 0 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnBiasIdaci = 0
            .bfEnVmi = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 1 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 0
            .bfEnBiasIdaci = 1
            .bfEnVmi = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 2 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnBiasIdaci = 0
            .bfEnVmi = 0
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    ElseIf piTxFilterSet = 3 Then
        With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
            .bfEnFilterbypass = 1
            .bfEnBiasIdaci = 1
            .bfEnVmi = 1
            Call Ctrx.WriteFixedRegister(.Self)
        End With
    End If
    
    If piTxFilterSet < 4 Then
        Call Dopp.Wait(WAIT_ISD_PA_SETTLE)
        Call Dopp.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Strobe(mdDcviStrobeWait)
    End If
    With CallByName(IfxRegMap, "NewRegTxTxrf" & piTxChan & "PsCf", VbCallType.VbMethod)
        .bfEnFilterbypass = 0
        .bfEnBiasIdaci = 0
        .bfEnVmi = 0
        Call Ctrx.WriteFixedRegister(.Self)
    End With
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Function reads the current measurements
'''
'''@details
''' for all different configurations the data is read from the meter and then post-processed
''' logs all current readouts and then the post-processed data
''' FBP0=(TxrfTxLeakageSetting1(1)-TxrfTxLeakageSetting1(0))
''' FBP1=(TxrfTxLeakageSetting1(3)-TxrfTxLeakageSetting1(2))
''' ratio=FBP0/FB1
'''
'''@log{2025-04-08, lef, Initial Version\, @JIRA{CTRXTE-5731}}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Private Sub TxgeFilterCurrentReadoutPSLog2()
    Const N_MEAS As Long = 4
    Const N_TXCH As Long = 8
    Const N_COMBINATIONS As Long = 5
    Const N_LEAK_TXPS As Long = N_COMBINATIONS * N_MEAS
    Dim liIdx As Long
    Dim liLeakAllSize As Long
    Dim lpldLeakTX(0 To 7) As New PinListData
    Dim lwLeakTX(0 To 7) As New DSPWave
    Dim lwLeakTXPS As New DSPWave
    
    Dim lwDeltaEn(0 To 7) As New DSPWave
    Dim lwDeltaBp(0 To 7) As New DSPWave
    Dim lwDeltaRatio(0 To 7) As New DSPWave
    
    Dim lwDeltaEnAll As New DSPWave
    Dim lwDeltaBpAll As New DSPWave
    Dim lwDeltaRatioAll As New DSPWave
    Dim liCtr As Long
    For Each Site In TheExec.Sites
        For liCtr = 0 To 7
            lpldLeakTX(liCtr) = TheHdw.Dcvi.Pins("VDD1V0TX_MDCVI").Meter.Read(tlNoStrobe, N_LEAK_TXPS, _
                , tlDCVIMeterReadingFormatArray)
            lwLeakTX(liCtr) = IfxMath.ConvertTo.DSPWave(lpldLeakTX(liCtr).Pins("VDD1V0TX_MDCVI"))
        Next liCtr
    Next Site
    
    If TheExec.TesterMode = testModeOffline Then
        For Each Site In TheExec.Sites
            For liCtr = 0 To 7
                lwLeakTX(liCtr).FileImport (".\Program\Tx\Waves\TxrfCurLeakMirrorTXPS_" & CStr(liCtr + 1) & ".txt")
            Next liCtr
        Next Site
    End If
    
    liLeakAllSize = N_LEAK_TXPS
    For Each Site In TheExec.Sites
        lwLeakTXPS = lwLeakTX(0).Concatenate(lwLeakTX(1)).Concatenate(lwLeakTX(2)).Concatenate(lwLeakTX(3)) _
            .Concatenate(lwLeakTX(4)).Concatenate(lwLeakTX(5)).Concatenate(lwLeakTX(6)).Concatenate(lwLeakTX(7))
    Next Site
    
    For liIdx = 0 To (N_MEAS * N_COMBINATIONS * N_TXCH) - 1
        Call IfxEnv.Datalog(lwLeakTXPS.Element(liIdx))
    Next liIdx
    
    For Each Site In TheExec.Sites
        For liCtr = 0 To 7
            Call CalcLeakMirrorsTXPS(lwLeakTX(liCtr), liLeakAllSize, lwDeltaEn(liCtr), lwDeltaBp(liCtr), _
                lwDeltaRatio(liCtr))
        Next liCtr
    Next Site
    
    For Each Site In TheExec.Sites
        lwDeltaEnAll = lwDeltaEn(0).Concatenate(lwDeltaEn(1)).Concatenate(lwDeltaEn(2)).Concatenate(lwDeltaEn(3)) _
            .Concatenate(lwDeltaEn(4)).Concatenate(lwDeltaEn(5)).Concatenate(lwDeltaEn(6)).Concatenate(lwDeltaEn(7))
        lwDeltaBpAll = lwDeltaBp(0).Concatenate(lwDeltaBp(1)).Concatenate(lwDeltaBp(2)).Concatenate(lwDeltaBp(3)) _
            .Concatenate(lwDeltaBp(4)).Concatenate(lwDeltaBp(5)).Concatenate(lwDeltaBp(6)).Concatenate(lwDeltaBp(7))
        lwDeltaRatioAll = lwDeltaRatio(0).Concatenate(lwDeltaRatio(1)).Concatenate(lwDeltaRatio(2)) _
            .Concatenate(lwDeltaRatio(3)).Concatenate(lwDeltaRatio(4)).Concatenate(lwDeltaRatio(5)) _
            .Concatenate(lwDeltaRatio(6)).Concatenate(lwDeltaRatio(7))
    Next Site
    
    For liIdx = 0 To (N_TXCH * N_COMBINATIONS) - 1
        Call IfxEnv.Datalog(lwDeltaEnAll.Element(liIdx))
        Call IfxEnv.Datalog(lwDeltaBpAll.Element(liIdx))
        Call IfxEnv.Datalog(lwDeltaRatioAll.Element(liIdx))
    Next liIdx
End Sub

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Calculate the difference values for CurWFiltEn and CurWFiltBp then calculate the ratios.
'''@details The Corrected Value = VDD + Offset. And the Offset = VDD - PreCalRead.
'''So for or calculation this gives us: Corrected value = VDD + VDD - PreCalRead
'''
'''@param[in] pwRawValues The measured current leakage values at the supply DCVIs
'''@param[in] pdRawSize The size of the RawValues array which must be divisible by 4.
'''@param[out] pwDeltaEn The difference array for WFiltEn current measurements
'''@param[out] pwDeltaBp The difference array for WFiltBp current measurements
'''@param[out] pwDeltaRatio The calculated ratio of the En value divided by the Bp value
'''
'''@log{2025-04-04, lef, @JIRA{CTRXTE-5684}: initial version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function CalcLeakMirrorsTXPS( _
    ByVal pwRawValues As DSPWave, _
    ByVal piRawSize As Long, _
    ByRef pwDeltaEn As DSPWave, _
    ByRef pwDeltaBp As DSPWave, _
    ByRef pwDeltaRatio As DSPWave) As Long
    
    Const N_MEAS As Long = 5
    Const NON_ZERO As Double = 5
    
    Dim lwCurWFiltEnA As New DSPWave
    Dim lwCurWFiltEnB As New DSPWave
    Dim lwCurWFiltBpA As New DSPWave
    Dim lwCurWFiltBpB As New DSPWave
    Dim lwZeroValues As New DSPWave
    Dim lwNonZero As New DSPWave
    
    ''' we expect the DspWave array size to match the RawSize parameter and this must be divisible by 4
    ''' - the RawValues contain multiple sets of 4 current measurements each
    If pwRawValues.SampleSize <> piRawSize Then
        'Measurement error -> likely a failure in the measured values which fails the test
        CalcLeakMirrorsTXPS = TL_ERROR
        Exit Function
    End If
    
    ''' separate out each of the 4 measurement types
    ''' - The first 2 relate to WFiltEn current measurements
    ''' - The second 2 relate to WFiltBp current measurements
    lwCurWFiltEnA = pwRawValues.Select(0, 1, N_MEAS)
    lwCurWFiltEnB = pwRawValues.Select(5, 1, N_MEAS)
    lwCurWFiltBpA = pwRawValues.Select(10, 1, N_MEAS)
    lwCurWFiltBpB = pwRawValues.Select(15, 1, N_MEAS)
    
    ''' Calculate the En and Bp current deltas as (B-A) and return these two results arrays
    Call pwDeltaEn.CreateConstant(0, N_MEAS, DspDouble)
    Call pwDeltaBp.CreateConstant(0, N_MEAS, DspDouble)
    pwDeltaEn = lwCurWFiltEnB.Subtract(lwCurWFiltEnA)
    pwDeltaBp = lwCurWFiltBpB.Subtract(lwCurWFiltBpA)
    
    ''' Calculate the ratio CurWFiltEn divided by CurWFiltBp and return this results array also
    ''' - includes check and replacement for elements with value 0 to avoid math error
    lwZeroValues = pwDeltaBp.FindIndices(EqualTo, 0)
    Call lwNonZero.CreateConstant(NON_ZERO, N_MEAS, DspDouble)
    Call pwDeltaBp.ReplaceElements(lwZeroValues, lwNonZero)
    
    Call pwDeltaRatio.CreateConstant(0, N_MEAS, DspDouble)
    pwDeltaRatio = pwDeltaEn.Divide(pwDeltaBp)
    
    CalcLeakMirrorsTXPS = TL_SUCCESS
End Function

'<AUTOCOMMENT:0.4:FUNCTION
'''@brief Calculate the difference values for CurWFiltEn and CurWFiltBp then calculate the ratios.
'''@details The Corrected Value = VDD + Offset. And the Offset = VDD - PreCalRead.
'''So for or calculation this gives us: Corrected value = VDD + VDD - PreCalRead
'''
'''@param[in] pwRawValues The measured current leakage values at the supply DCVIs
'''@param[in] pdRawSize The size of the RawValues array which must be divisible by 4.
'''@param[out] pwDeltaEn The difference array for WFiltEn current measurements
'''@param[out] pwDeltaBp The difference array for WFiltBp current measurements
'''@param[out] pwDeltaRatio The calculated ratio of the En value divided by the Bp value
'''
'''@log{2025-04-04, lef, @JIRA{CTRXTE-5684}: initial version}
'''@log{2026-01-20, Schilling Johannes, Initial implementation for 8188\, @JIRA{RSIPPTE-542}}
'>AUTOCOMMENT:0.4:FUNCTION
Public Function CalcLeakMirrorsTXPA( _
    ByVal pwRawValues As DSPWave, _
    ByVal piRawSize As Long, _
    ByRef pwDeltaEn As DSPWave, _
    ByRef pwDeltaBp As DSPWave, _
    ByRef pwDeltaRatio As DSPWave) As Long
    
    Const N_MEAS As Long = 4
    Const NON_ZERO As Double = 5
    Dim liResultSize As Long
    Dim lwCurWFiltEnA As New DSPWave
    Dim lwCurWFiltEnB As New DSPWave
    Dim lwCurWFiltBpA As New DSPWave
    Dim lwCurWFiltBpB As New DSPWave
    Dim lwZeroValues As New DSPWave
    Dim lwNonZero As New DSPWave
    
    ''' we expect the DspWave array size to match the RawSize parameter and this must be divisible by 4
    ''' - the RawValues contain multiple sets of 4 current measurements each
    If pwRawValues.SampleSize <> piRawSize Then
        'Measurement error -> likely a failure in the measured values which fails the test
        CalcLeakMirrorsTXPA = TL_ERROR
        Exit Function
    End If
    
    ''' separate out each of the 4 measurement types
    ''' - The first 2 relate to WFiltEn current measurements
    ''' - The second 2 relate to WFiltBp current measurements
    lwCurWFiltEnA = pwRawValues.Select(0, 1, N_MEAS)
    lwCurWFiltEnB = pwRawValues.Select(4, 1, N_MEAS)
    lwCurWFiltBpA = pwRawValues.Select(8, 1, N_MEAS)
    lwCurWFiltBpB = pwRawValues.Select(12, 1, N_MEAS)
    
    ''' Calculate the En and Bp current deltas as (B-A) and return these two results arrays
    liResultSize = piRawSize / N_MEAS
    Call pwDeltaEn.CreateConstant(0, liResultSize, DspDouble)
    Call pwDeltaBp.CreateConstant(0, liResultSize, DspDouble)
    pwDeltaEn = lwCurWFiltEnB.Subtract(lwCurWFiltEnA)
    pwDeltaBp = lwCurWFiltBpB.Subtract(lwCurWFiltBpA)
    
    ''' Calculate the ratio CurWFiltEn divided by CurWFiltBp and return this results array also
    ''' - includes check and replacement for elements with value 0 to avoid math error
    lwZeroValues = pwDeltaBp.FindIndices(EqualTo, 0)
    Call lwNonZero.CreateConstant(NON_ZERO, liResultSize, DspDouble)
    Call pwDeltaBp.ReplaceElements(lwZeroValues, lwNonZero)
    
    Call pwDeltaRatio.CreateConstant(0, liResultSize, DspDouble)
    pwDeltaRatio = pwDeltaEn.Divide(pwDeltaBp)
    
    CalcLeakMirrorsTXPA = TL_SUCCESS
End Function
