Attribute VB_Name = "apmu_drv"
Option Explicit

' tl_APMUReset() - Handle APMU reset
Declare Function tl_APMUReset Lib "apmu_drv.dll" Alias "_tl_APMUReset@4" (ByVal resetType As Long) As Long

' tl_APMU_setErrorOutput() - redirect error output for APMU errors
Declare Function tl_APMU_setErrorOutput Lib "apmu_drv.dll" Alias "_tl_APMU_setErrorOutput@12" (ByVal where As Long, ByRef fileName As String, ByRef mode As String) As Long

' tl_APMU_SetMode() - select APMU operation mode
Declare Function tl_APMU_SetMode Lib "apmu_drv.dll" Alias "_tl_APMU_SetMode@16" (ByRef Channels() As Long, ByVal opmode As Long, ByVal vrng As Long, ByVal irng As Long) As Long

' tl_APMU_GetMode() - Read and return mode and ranges
Declare Function tl_APMU_GetMode Lib "apmu_drv.dll" Alias "_tl_APMU_GetMode@16" (ByRef Channels() As Long, ByRef modes() As Long, ByRef vrngs() As Long, ByRef irngs() As Long) As Long

' tl_APMU_SetForceVoltage() - set force voltage
Declare Function tl_APMU_SetForceVoltage Lib "apmu_drv.dll" Alias "_tl_APMU_SetForceVoltage@16" (ByRef Channels() As Long, ByVal vrng As Long, ByVal vval As Double) As Long

' tl_APMU_GetForceVoltage() - Read and return force voltage
Declare Function tl_APMU_GetForceVoltage Lib "apmu_drv.dll" Alias "_tl_APMU_GetForceVoltage@12" (ByRef Channels() As Long, ByRef vrngs() As Long, ByRef vvals() As Double) As Long

' tl_APMU_SetForceCurrent() - set force current
Declare Function tl_APMU_SetForceCurrent Lib "apmu_drv.dll" Alias "_tl_APMU_SetForceCurrent@16" (ByRef Channels() As Long, ByVal irng As Long, ByVal IVal As Double) As Long

' tl_APMU_GetForceCurrent() - Read and return force current
Declare Function tl_APMU_GetForceCurrent Lib "apmu_drv.dll" Alias "_tl_APMU_GetForceCurrent@12" (ByRef Channels() As Long, ByRef irngs() As Long, ByRef ivals() As Double) As Long

' tl_APMU_SetClampVoltage() - set clamp voltage
Declare Function tl_APMU_SetClampVoltage Lib "apmu_drv.dll" Alias "_tl_APMU_SetClampVoltage@16" (ByRef Channels() As Long, ByVal vrng As Long, ByVal vclamp As Double) As Long

' tl_APMU_GetClampVoltage() - Read and return clamp voltage
Declare Function tl_APMU_GetClampVoltage Lib "apmu_drv.dll" Alias "_tl_APMU_GetClampVoltage@12" (ByRef Channels() As Long, ByRef vrngs() As Long, ByRef vclamps() As Double) As Long

' tl_APMU_SetClampCurrent() - set clamp current
Declare Function tl_APMU_SetClampCurrent Lib "apmu_drv.dll" Alias "_tl_APMU_SetClampCurrent@16" (ByRef Channels() As Long, ByVal irng As Long, ByVal iclamp As Double) As Long

' tl_APMU_GetClampCurrent() - Read and return clamp current
Declare Function tl_APMU_GetClampCurrent Lib "apmu_drv.dll" Alias "_tl_APMU_GetClampCurrent@12" (ByRef Channels() As Long, ByRef irngs() As Long, ByRef iclamps() As Double) As Long

' tl_APMU_SetGate() - switch gate
Declare Function tl_APMU_SetGate Lib "apmu_drv.dll" Alias "_tl_APMU_SetGate@8" (ByRef Channels() As Long, ByVal gateonoff As Long) As Long

' tl_APMU_GetGate() - Read and return gate state
Declare Function tl_APMU_GetGate Lib "apmu_drv.dll" Alias "_tl_APMU_GetGate@8" (ByRef Channels() As Long, ByRef gates() As Long) As Long

' tl_APMU_SetConnect() - connect the specified channels
Declare Function tl_APMU_SetConnect Lib "apmu_drv.dll" Alias "_tl_APMU_SetConnect@8" (ByRef Channels() As Long, ByVal relayState As Long) As Long

' tl_APMU_GetConnect() - Read and return relay state
Declare Function tl_APMU_GetConnect Lib "apmu_drv.dll" Alias "_tl_APMU_GetConnect@8" (ByRef Channels() As Long, ByRef relays() As Long) As Long

' tl_APMU_Measure() - take measurements from the specified channels
Declare Function tl_APMU_Measure Lib "apmu_drv.dll" Alias "_tl_APMU_Measure@12" (ByRef Channels() As Long, ByVal nSamples As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_MeasureTime() - take measurements from the specified channels
Declare Function tl_APMU_MeasureTime Lib "apmu_drv.dll" Alias "_tl_APMU_MeasureTime@16" (ByRef Channels() As Long, ByVal mTime As Double, ByRef measuredValues() As Double) As Long

' tl_APMU_MeasureDiffVMeter() - take measurements from differential voltage meter
Declare Function tl_APMU_MeasureDiffVMeter Lib "apmu_drv.dll" Alias "_tl_APMU_MeasureDiffVMeter@24" (ByRef channels_h() As Long, ByRef channels_l() As Long, ByVal nSamples As Long, ByVal dmRng As Long, ByVal lpfEnable As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_MeasureDiffVMeterTime() - take measurements from differential voltage meter
Declare Function tl_APMU_MeasureDiffVMeterTime Lib "apmu_drv.dll" Alias "_tl_APMU_MeasureDiffVMeterTime@28" (ByRef channels_h() As Long, ByRef channels_l() As Long, ByVal mTime As Double, ByVal dmRng As Long, ByVal lpfEnable As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_SetLPF() - set LPF for meter
Declare Function tl_APMU_SetLPF Lib "apmu_drv.dll" Alias "_tl_APMU_SetLPF@8" (ByRef Channels() As Long, ByVal lpfEnable As Long) As Long

' tl_APMU_GetLPF() - Read and return LPF state
Declare Function tl_APMU_GetLPF Lib "apmu_drv.dll" Alias "_tl_APMU_GetLPF@8" (ByRef Channels() As Long, ByRef lpfs() As Long) As Long

' tl_APMU_SetGang() - set gang mode
Declare Function tl_APMU_SetGang Lib "apmu_drv.dll" Alias "_tl_APMU_SetGang@8" (ByRef Channels() As Long, ByVal gangOnoff As Long) As Long

' tl_APMU_GetGang() - Read and return Gang mode
Declare Function tl_APMU_GetGang Lib "apmu_drv.dll" Alias "_tl_APMU_GetGang@8" (ByRef Channels() As Long, ByRef gangs() As Long) As Long

' tl_APMU_SetExternalSense() - set External Sense
Declare Function tl_APMU_SetExternalSense Lib "apmu_drv.dll" Alias "_tl_APMU_SetExternalSense@8" (ByRef Channels() As Long, ByVal extsenseOnoff As Long) As Long

' tl_APMU_GetExternalSense() - Read and return External Sense
Declare Function tl_APMU_GetExternalSense Lib "apmu_drv.dll" Alias "_tl_APMU_GetExternalSense@8" (ByRef Channels() As Long, ByRef extsenses() As Long) As Long

' tl_APMU_SetModePins() - select APMU operation mode
Declare Function tl_APMU_SetModePins Lib "apmu_drv.dll" Alias "_tl_APMU_SetModePins@16" (ByRef p_name As String, ByVal opmode As Long, ByVal vrng As Long, ByVal irng As Long) As Long

' tl_APMU_GetModePins() - Read and return mode and ranges
Declare Function tl_APMU_GetModePins Lib "apmu_drv.dll" Alias "_tl_APMU_GetModePins@16" (ByRef p_name As String, ByRef modes() As Long, ByRef vrngs() As Long, ByRef irngs() As Long) As Long

' tl_APMU_SetForceVoltagePins() - set force voltage
Declare Function tl_APMU_SetForceVoltagePins Lib "apmu_drv.dll" Alias "_tl_APMU_SetForceVoltagePins@16" (ByRef p_name As String, ByVal vrng As Long, ByVal vval As Double) As Long

' tl_APMU_GetForceVoltagePins() - Read and return force voltage
Declare Function tl_APMU_GetForceVoltagePins Lib "apmu_drv.dll" Alias "_tl_APMU_GetForceVoltagePins@12" (ByRef p_name As String, ByRef vrngs() As Long, ByRef vvals() As Double) As Long

' tl_APMU_SelectMeasureCurrentRange() - select current range
Declare Function tl_APMU_SelectMeasureCurrentRange Lib "apmu_drv.dll" Alias "_tl_APMU_SelectMeasureCurrentRange@12" (ByVal IVal As Double, ByRef irng As Long) As Long

' tl_APMU_SelectMeasureCurrentRangePin() - select current range
Declare Function tl_APMU_SelectMeasureCurrentRangePin Lib "apmu_drv.dll" Alias "_tl_APMU_SelectMeasureCurrentRangePin@16" (ByRef p_name As String, ByVal IVal As Double, ByRef irng As Long) As Long

' tl_APMU_SelectForceCurrentRange() - select current range
Declare Function tl_APMU_SelectForceCurrentRange Lib "apmu_drv.dll" Alias "_tl_APMU_SelectForceCurrentRange@12" (ByVal IVal As Double, ByRef irng As Long) As Long

' tl_APMU_SelectForceCurrentRangePin() - select current range
Declare Function tl_APMU_SelectForceCurrentRangePin Lib "apmu_drv.dll" Alias "_tl_APMU_SelectForceCurrentRangePin@16" (ByRef p_name As String, ByVal IVal As Double, ByRef irng As Long) As Long

' tl_APMU_SelectVoltageRange() - select voltage range
Declare Function tl_APMU_SelectVoltageRange Lib "apmu_drv.dll" Alias "_tl_APMU_SelectVoltageRange@12" (ByVal vval As Double, ByRef vrng As Long) As Long

' tl_APMU_SetForceCurrentPins() - set force current
Declare Function tl_APMU_SetForceCurrentPins Lib "apmu_drv.dll" Alias "_tl_APMU_SetForceCurrentPins@16" (ByRef p_name As String, ByVal irng As Long, ByVal IVal As Double) As Long

' tl_APMU_GetForceCurrentPins() - Read and return force current
Declare Function tl_APMU_GetForceCurrentPins Lib "apmu_drv.dll" Alias "_tl_APMU_GetForceCurrentPins@12" (ByRef p_name As String, ByRef irngs() As Long, ByRef ivals() As Double) As Long

' tl_APMU_SetClampVoltagePins() - set clamp voltage
Declare Function tl_APMU_SetClampVoltagePins Lib "apmu_drv.dll" Alias "_tl_APMU_SetClampVoltagePins@16" (ByRef p_name As String, ByVal vrng As Long, ByVal vclamp As Double) As Long

' tl_APMU_GetClampVoltagePins() - Read and return clamp voltage
Declare Function tl_APMU_GetClampVoltagePins Lib "apmu_drv.dll" Alias "_tl_APMU_GetClampVoltagePins@12" (ByRef p_name As String, ByRef vrngs() As Long, ByRef vclamps() As Double) As Long

' tl_APMU_SetClampCurrentPins() - set clamp current
Declare Function tl_APMU_SetClampCurrentPins Lib "apmu_drv.dll" Alias "_tl_APMU_SetClampCurrentPins@16" (ByRef p_name As String, ByVal irng As Long, ByVal iclamp As Double) As Long

' tl_APMU_GetClampCurrentPins() - Read and return clamp current
Declare Function tl_APMU_GetClampCurrentPins Lib "apmu_drv.dll" Alias "_tl_APMU_GetClampCurrentPins@12" (ByRef p_name As String, ByRef irngs() As Long, ByRef iclamps() As Double) As Long

' tl_APMU_SetGatePins() - switch gate
Declare Function tl_APMU_SetGatePins Lib "apmu_drv.dll" Alias "_tl_APMU_SetGatePins@8" (ByRef p_name As String, ByVal gateonoff As Long) As Long

' tl_APMU_GetGatePins() - Read and return gate state
Declare Function tl_APMU_GetGatePins Lib "apmu_drv.dll" Alias "_tl_APMU_GetGatePins@8" (ByRef p_name As String, ByRef gates() As Long) As Long

' tl_APMU_SetConnectPins() - connect the specified channels
Declare Function tl_APMU_SetConnectPins Lib "apmu_drv.dll" Alias "_tl_APMU_SetConnectPins@8" (ByRef p_name As String, ByVal relayState As Long) As Long

' tl_APMU_GetConnectPins() - Read and return relay state
Declare Function tl_APMU_GetConnectPins Lib "apmu_drv.dll" Alias "_tl_APMU_GetConnectPins@8" (ByRef p_name As String, ByRef relays() As Long) As Long

' tl_APMU_MeasurePins() - take measurements from the specified channels
Declare Function tl_APMU_MeasurePins Lib "apmu_drv.dll" Alias "_tl_APMU_MeasurePins@12" (ByRef p_name As String, ByVal nSamples As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_MeasureTimePins() - take measurements from the specified channels
Declare Function tl_APMU_MeasureTimePins Lib "apmu_drv.dll" Alias "_tl_APMU_MeasureTimePins@16" (ByRef p_name As String, ByVal mTime As Double, ByRef measuredValues() As Double) As Long

' tl_APMU_MeasureDiffVMeterPins() - take measurements from differential voltage meter
Declare Function tl_APMU_MeasureDiffVMeterPins Lib "apmu_drv.dll" Alias "_tl_APMU_MeasureDiffVMeterPins@24" (ByRef p_name_h As String, ByRef p_name_l As String, ByVal nSamples As Long, ByVal dmRng As Long, ByVal lpfEnable As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_MeasureDiffVMeterTimePins() - take measurements from differential voltage meter
Declare Function tl_APMU_MeasureDiffVMeterTimePins Lib "apmu_drv.dll" Alias "_tl_APMU_MeasureDiffVMeterTimePins@28" (ByRef p_name_h As String, ByRef p_name_l As String, ByVal mTime As Double, ByVal dmRng As Long, ByVal lpfEnable As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_SetLPFPins() - set LPF for meter
Declare Function tl_APMU_SetLPFPins Lib "apmu_drv.dll" Alias "_tl_APMU_SetLPFPins@8" (ByRef p_name As String, ByVal lpfEnable As Long) As Long

' tl_APMU_GetLPFPins() - Read and return LPF state
Declare Function tl_APMU_GetLPFPins Lib "apmu_drv.dll" Alias "_tl_APMU_GetLPFPins@8" (ByRef p_name As String, ByRef lpfs() As Long) As Long

' tl_APMU_SetGangPins() - set gang mode
Declare Function tl_APMU_SetGangPins Lib "apmu_drv.dll" Alias "_tl_APMU_SetGangPins@8" (ByRef p_name As String, ByVal gangOnoff As Long) As Long

' tl_APMU_GetGangPins() - Read and return Gang mode
Declare Function tl_APMU_GetGangPins Lib "apmu_drv.dll" Alias "_tl_APMU_GetGangPins@8" (ByRef p_name As String, ByRef gangs() As Long) As Long

' tl_APMU_SetExternalSensePins() - set External Sense
Declare Function tl_APMU_SetExternalSensePins Lib "apmu_drv.dll" Alias "_tl_APMU_SetExternalSensePins@8" (ByRef p_name As String, ByVal extsenseOnoff As Long) As Long

' tl_APMU_GetExternalSensePins() - Read and return External Sense
Declare Function tl_APMU_GetExternalSensePins Lib "apmu_drv.dll" Alias "_tl_APMU_GetExternalSensePins@8" (ByRef p_name As String, ByRef extsenses() As Long) As Long

' tl_APMU_SetMeasureWaitTime() - set wait time for measure
Declare Function tl_APMU_SetMeasureWaitTime Lib "apmu_drv.dll" Alias "_tl_APMU_SetMeasureWaitTime@8" (ByVal wTime As Double) As Long

' tl_APMU_GetMeasureWaitTime() - Read and return wait time
Declare Function tl_APMU_GetMeasureWaitTime Lib "apmu_drv.dll" Alias "_tl_APMU_GetMeasureWaitTime@4" (ByRef wTime As Double) As Long

' tl_APMU_SetMeasurePeriod() - set period time for measure
Declare Function tl_APMU_SetMeasurePeriod Lib "apmu_drv.dll" Alias "_tl_APMU_SetMeasurePeriod@8" (ByVal mPeriod As Double) As Long

' tl_APMU_GetMeasurePeriod() - Read and return period time
Declare Function tl_APMU_GetMeasurePeriod Lib "apmu_drv.dll" Alias "_tl_APMU_GetMeasurePeriod@4" (ByRef mPeriod As Double) As Long

' tl_APMU_GetMeasureCount() - Return measurement count
Declare Function tl_APMU_GetMeasureCount Lib "apmu_drv.dll" Alias "_tl_APMU_GetMeasureCount@4" (ByRef mCnt As Long) As Long

' tl_APMU_CKRConnectForce() - connect force line to CALBUS Force
Declare Function tl_APMU_CKRConnectForce Lib "apmu_drv.dll" Alias "_tl_APMU_CKRConnectForce@4" (ByVal ChannelNumber As Long) As Long

' tl_APMU_CKRDisconnectForce() - disconnect force line from CALBUS Force
Declare Function tl_APMU_CKRDisconnectForce Lib "apmu_drv.dll" Alias "_tl_APMU_CKRDisconnectForce@4" (ByVal ChannelNumber As Long) As Long

' tl_APMU_CKRConnectSense() - connect to internal buses
Declare Function tl_APMU_CKRConnectSense Lib "apmu_drv.dll" Alias "_tl_APMU_CKRConnectSense@8" (ByRef Channels() As Long, ByVal relayval As Long) As Long

' tl_APMU_CKRConnectRMW() - connect to internal buses
Declare Function tl_APMU_CKRConnectRMW Lib "apmu_drv.dll" Alias "_tl_APMU_CKRConnectRMW@12" (ByRef Channels() As Long, ByVal relayval As Long, ByVal readmask) As Long

' tl_APMU_CKRConnect() - connect to internal buses
Declare Function tl_APMU_CKRConnect Lib "apmu_drv.dll" Alias "_tl_APMU_CKRConnect@8" (ByRef Channels() As Long, ByVal relayReg As Long) As Long

' tl_APMU_CKRKelvin() - connect/disconnect kelvin relay
Declare Function tl_APMU_CKRKelvin Lib "apmu_drv.dll" Alias "_tl_APMU_CKRKelvin@8" (ByRef Channels() As Long, ByVal kelvinRly As Long) As Long

' tl_APMU_CKRRelayControl() - enable/disable relay connection
Declare Function tl_APMU_CKRRelayControl Lib "apmu_drv.dll" Alias "_tl_APMU_CKRRelayControl@4" (ByVal ctrlOnoff As Long) As Long

' tl_APMU_DMInput() - select input for differential voltage meter
Declare Function tl_APMU_DMInput Lib "apmu_drv.dll" Alias "_tl_APMU_DMInput@8" (ByRef meters() As Long, ByVal dmSrc As Long) As Long

' tl_APMU_DMRange() - select input range for differential voltage meter
Declare Function tl_APMU_DMRange Lib "apmu_drv.dll" Alias "_tl_APMU_DMRange@8" (ByRef meters() As Long, ByVal dmRng As Long) As Long

' tl_APMU_DMMeasure() - take measurements from differential voltage meter
Declare Function tl_APMU_DMMeasure Lib "apmu_drv.dll" Alias "_tl_APMU_DMMeasure@12" (ByRef meters() As Long, ByVal nSamples As Long, ByRef measuredValues() As Double) As Long

' tl_APMU_CalSetR() - select R Ref.
Declare Function tl_APMU_CalSetR Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetR@8" (ByVal bdnum As Long, ByVal rref As Long) As Long

' tl_APMU_CalSetRBds() - select R Ref.
Declare Function tl_APMU_CalSetRBds Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetRBds@8" (ByRef boards As Long, ByVal rref As Long) As Long

' tl_APMU_CalSetRS() - select R Ref.
Declare Function tl_APMU_CalSetRS Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetRS@8" (ByVal bdnum As Long, ByVal rref As Long) As Long

' tl_APMU_CalSetRSBds() - select R Ref.
Declare Function tl_APMU_CalSetRSBds Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetRSBds@8" (ByRef boards As Long, ByVal rref As Long) As Long

' tl_APMU_CalSetV() - select V Ref.
Declare Function tl_APMU_CalSetV Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetV@8" (ByVal bdnum As Long, ByVal vref As Long) As Long

' tl_APMU_CalSetVBds() - select V Ref.
Declare Function tl_APMU_CalSetVBds Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetVBds@8" (ByRef boards As Long, ByVal vref As Long) As Long

' tl_APMU_CalSetVVal() - select V Ref.
Declare Function tl_APMU_CalSetVVal Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetVVal@12" (ByVal bdnum As Long, ByVal vref As Double) As Long

' tl_APMU_CalSetVValBds() - select V Ref.
Declare Function tl_APMU_CalSetVValBds Lib "apmu_drv.dll" Alias "_tl_APMU_CalSetVValBds@12" (ByRef boards As Long, ByVal vref As Double) As Long

' tl_APMU_CalConnectRef() - connect reference to VM bus
Declare Function tl_APMU_CalConnectRef Lib "apmu_drv.dll" Alias "_tl_APMU_CalConnectRef@12" (ByVal bdnum As Long, ByVal calSrc_h As Long, ByVal calSrc_l As Long) As Long

' tl_APMU_CalConnectRefBds() - connect reference to VM bus
Declare Function tl_APMU_CalConnectRefBds Lib "apmu_drv.dll" Alias "_tl_APMU_CalConnectRefBds@12" (ByRef boards As Long, ByVal calSrc_h As Long, ByVal calSrc_l As Long) As Long

' tl_APMU_CalVMBusReg() - select V Ref.
Declare Function tl_APMU_CalVMBusReg Lib "apmu_drv.dll" Alias "_tl_APMU_CalVMBusReg@8" (ByVal bdnum As Long, ByVal regval As Long) As Long

' tl_APMU_CalVMBusRegBds() - select V Ref.
Declare Function tl_APMU_CalVMBusRegBds Lib "apmu_drv.dll" Alias "_tl_APMU_CalVMBusRegBds@8" (ByRef boards As Long, ByVal regval As Long) As Long

' tl_APMU_CalParamCtrl() - calibration parameter control
Declare Function tl_APMU_CalParamCtrl Lib "apmu_drv.dll" Alias "_tl_APMU_CalParamCtrl@4" (ByVal which As Long) As Long

' tl_APMU_SetAlarm() - Enable/Disable alarm detection
Declare Function tl_APMU_SetAlarm Lib "apmu_drv.dll" Alias "_tl_APMU_SetAlarm@8" (ByRef Channels() As Long, ByVal alarmOnoff As Long) As Long

' tl_APMU_GetAlarm() - Read and return Alarm detection
Declare Function tl_APMU_GetAlarm Lib "apmu_drv.dll" Alias "_tl_APMU_GetAlarm@8" (ByRef Channels() As Long, ByRef alarms() As Long) As Long

' tl_APMU_SetAlarmPins() - Enable/Disable alarm detection
Declare Function tl_APMU_SetAlarmPins Lib "apmu_drv.dll" Alias "_tl_APMU_SetAlarmPins@8" (ByRef p_name As String, ByVal alarmOnoff As Long) As Long

' tl_APMU_GetAlarmPins() - Read and return Alarm detection
Declare Function tl_APMU_GetAlarmPins Lib "apmu_drv.dll" Alias "_tl_APMU_GetAlarmPins@8" (ByRef p_name As String, ByRef alarms() As Long) As Long

' tl_APMU_AlarmCheck() - check overload alarm
Declare Function tl_APMU_AlarmCheck Lib "apmu_drv.dll" Alias "_tl_APMU_AlarmCheck@0" () As Long

' tl_APMU_SetUtilBitBoard() - control utility bit
Declare Function tl_APMU_SetUtilBitBoard Lib "apmu_drv.dll" Alias "_tl_APMU_SetUtilBitBoard@4" (ByRef boards() As Long) As Long

' tl_APMU_SetUtilBit() - control utility bit
Declare Function tl_APMU_SetUtilBit Lib "apmu_drv.dll" Alias "_tl_APMU_SetUtilBit@8" (ByRef bits() As Long, ByVal bitOnoff As Long) As Long

' tl_APMU_SetUtilBitOne() - control utility bit
Declare Function tl_APMU_SetUtilBitOne Lib "apmu_drv.dll" Alias "_tl_APMU_SetUtilBitOne@12" (ByVal bdnum As Long, ByVal bitnum As Long, ByVal bitOnoff As Long) As Long

' tl_APMU_GetUtilBitOne() - control utility bit
Declare Function tl_APMU_GetUtilBitOne Lib "apmu_drv.dll" Alias "_tl_APMU_GetUtilBitOne@12" (ByVal bdnum As Long, ByVal bitnum As Long, ByRef bitOnoff As Long) As Long

' tl_APMU_GetUtilBit() - control utility bit
Declare Function tl_APMU_GetUtilBit Lib "apmu_drv.dll" Alias "_tl_APMU_GetUtilBit@8" (ByVal bitnum As Long, ByRef bitOnoff As Long) As Long

' tl_APMU_SetFastMode() - Enable/Disable fast mode
Declare Function tl_APMU_SetFastMode Lib "apmu_drv.dll" Alias "_tl_APMU_SetFastMode@4" (ByVal fastmode As Long) As Long

' tl_APMU_SetAutoCalOn() - Enable auto calibration
Declare Function tl_APMU_SetAutoCalOn Lib "apmu_drv.dll" Alias "_tl_APMU_SetAutoCalOn@4" (ByVal interval As Long) As Long

' tl_APMU_SetAutoCalOff() - Disable auto calibration
Declare Function tl_APMU_SetAutoCalOff Lib "apmu_drv.dll" Alias "_tl_APMU_SetAutoCalOff@0" () As Long

' tl_APMU_CalExecute() - Execute calibration
Declare Function tl_APMU_CalExecute Lib "apmu_drv.dll" Alias "_tl_APMU_CalExecute@4" (ByVal flag As Long) As Long

' tl_APMU_GetCalStatus - get Cal status
Declare Function tl_APMU_GetCalStatus Lib "apmu_drv.dll" Alias "_tl_APMU_GetCalStatus@0" () As Long

' tl_APMU_GetRelayStatus - return F/S/ES relay status
Declare Function tl_APMU_GetRelayStatus Lib "apmu_drv.dll" Alias "_tl_APMU_GetRelayStatus@16" (ByVal ch As Long, ByRef f As Long, ByRef S As Long, ByRef exts As Long) As Long

' tl_APMU_SetIdPromSerial - set IdProm
Declare Function tl_APMU_SetIdPromSerial Lib "apmu_drv.dll" Alias "_tl_APMU_SetIdPromSerial@28" (ByVal bdtype As Long, ByVal bdnum As Long, ByVal subbdnum As Long, ByVal company As Long, ByVal board As Long, ByVal rev As Long, ByVal serial As Long) As Long

' tl_APMU_GetIdPromSerial - set IdProm
Declare Function tl_APMU_GetIdPromSerial Lib "apmu_drv.dll" Alias "_tl_APMU_GetIdPromSerial@28" (ByVal bdtype As Long, ByVal bdnum As Long, ByVal subbdnum As Long, ByRef company As Long, ByRef board As Long, ByRef rev As Long, ByRef serial As Long) As Long

' tl_APMU_SetIdPromRRefError - set IdProm
Declare Function tl_APMU_SetIdPromRRefError Lib "apmu_drv.dll" Alias "_tl_APMU_SetIdPromRRefError@16" (ByVal bdnum As Long, ByVal rref As Long, ByVal Err As Double) As Long

' tl_APMU_GetIdPromRRefError - get IdProm
Declare Function tl_APMU_GetIdPromRRefError Lib "apmu_drv.dll" Alias "_tl_APMU_GetIdPromRRefError@12" (ByVal bdnum As Long, ByVal rref As Long, ByRef Err As Double) As Long

' tl_APMU_SetIdPromVRefError - set IdProm
Declare Function tl_APMU_SetIdPromVRefError Lib "apmu_drv.dll" Alias "_tl_APMU_SetIdPromVRefError@16" (ByVal bdnum As Long, ByVal vref As Long, ByVal Err As Double) As Long

' tl_APMU_GetIdPromVRefError - get IdProm
Declare Function tl_APMU_GetIdPromVRefError Lib "apmu_drv.dll" Alias "_tl_APMU_GetIdPromVRefError@12" (ByVal bdnum As Long, ByVal vref As Long, ByRef Err As Double) As Long

' tl_APMU_SetIdPromRefDACError - set IdProm
Declare Function tl_APMU_SetIdPromRefDACError Lib "apmu_drv.dll" Alias "_tl_APMU_SetIdPromRefDACError@24" (ByVal bdnum As Long, ByVal rng As Long, ByVal gain As Double, ByVal offset As Double) As Long

' tl_APMU_GetIdPromRefDACError - get IdProm
Declare Function tl_APMU_GetIdPromRefDACError Lib "apmu_drv.dll" Alias "_tl_APMU_GetIdPromRefDACError@16" (ByVal bdnum As Long, ByVal rng As Long, ByRef gain As Double, ByRef offset As Double) As Long

' tl_APMU_SetIdPromData - set IdProm data
Declare Function tl_APMU_SetIdPromData Lib "apmu_drv.dll" Alias "_tl_APMU_SetIdPromData@20" (ByVal bdtype As Long, ByVal bdnum As Long, ByVal subbdnum As Long, ByVal addr As Long, ByVal Data As Long) As Long

' tl_APMU_GetIdPromData - get IdProm data
Declare Function tl_APMU_GetIdPromData Lib "apmu_drv.dll" Alias "_tl_APMU_GetIdPromData@20" (ByVal bdtype As Long, ByVal bdnum As Long, ByVal subbdnum As Long, ByVal addr As Long, ByRef Data As Long) As Long

' tl_APMU_GetTemperature - get Temperature
Declare Function tl_APMU_GetTemperature Lib "apmu_drv.dll" Alias "_tl_APMU_GetTemperature@4" (ByRef values() As Double) As Long

' tl_APMU_DumpCalParam - get calibration parameter
Declare Function tl_APMU_DumpCalParam Lib "apmu_drv.dll" Alias "_tl_APMU_DumpCalParam@8" (ByRef Channels() As Long, ByRef fileName As String) As Long


Public Const TL_APMU_V_UNKNOWN As Long = -1
Public Const TL_APMU_V_2V As Long = 0
Public Const TL_APMU_V_5V As Long = 1
Public Const TL_APMU_V_10V As Long = 2
Public Const TL_APMU_V_40V As Long = 3
Public Const TL_APMU_V_35V As Long = 4
Public Const TL_APMU_V_ITEM_COUNT As Long = 5

Public Const TL_APMU_I_UNKNOWN As Long = -1
Public Const TL_APMU_I_200UA As Long = 0
Public Const TL_APMU_I_1MA As Long = 1
Public Const TL_APMU_I_5MA As Long = 2
Public Const TL_APMU_I_50MA As Long = 3
Public Const TL_APMU_I_200NA As Long = 4
Public Const TL_APMU_I_2UA As Long = 5
Public Const TL_APMU_I_10UA As Long = 6
Public Const TL_APMU_I_40UA As Long = 7
Public Const TL_APMU_I_ITEM_COUNT As Long = 8

Public Const TL_APMU_DM_UNKNOWN As Long = -1
Public Const TL_APMU_DM_10V As Long = 0
Public Const TL_APMU_DM_100MV As Long = 1
Public Const TL_APMU_DM_1V As Long = 2
Public Const TL_APMU_DM_ITEM_COUNT As Long = 3

Public Const TL_APMU_DM_NORMAL As Long = 1
Public Const TL_APMU_DM_LPF As Long = 2
Public Const TL_APMU_DM_CHKA As Long = 4
Public Const TL_APMU_DM_CHKB As Long = 8

Public Const TL_APMU_MODE_UNKNOWN As Long = -1
Public Const TL_APMU_FVMI_MODE As Long = 0
Public Const TL_APMU_FIMV_MODE As Long = 1
Public Const TL_APMU_FVMV_MODE As Long = 2
Public Const TL_APMU_FIMI_MODE As Long = 3
Public Const TL_APMU_MV_MODE As Long = 4

Public Const TL_APMU_RELAY_NONE As Long = &H0
Public Const TL_APMU_RELAY_FORCE As Long = &H1
Public Const TL_APMU_RELAY_SENSE As Long = &H2
Public Const TL_APMU_CKR_VMH As Long = &H4
Public Const TL_APMU_CKR_VML As Long = &H8
Public Const TL_APMU_CKR_CALF As Long = &H10
Public Const TL_APMU_CKR_CALS As Long = &H20
Public Const TL_APMU_CKR_CALG As Long = &H40
Public Const TL_APMU_RELAY_EXT_SENSE As Long = &H80
Public Const TL_APMU_RELAY_GANG As Long = &H80

Public Const TL_APMU_RREF_NONE As Long = 0
Public Const TL_APMU_RREF_10 As Long = 1
Public Const TL_APMU_RREF_80 As Long = 2
Public Const TL_APMU_RREF_800 As Long = 4
Public Const TL_APMU_RREF_8K As Long = 8
Public Const TL_APMU_RREF_40K As Long = 16
Public Const TL_APMU_RREF_200K As Long = 32
Public Const TL_APMU_RREF_10M As Long = 64

Public Const TL_APMU_VREF_NONE As Long = 0
Public Const TL_APMU_VREF_P2 As Long = 20
Public Const TL_APMU_VREF_P5 As Long = 36
Public Const TL_APMU_VREF_P10 As Long = 68
Public Const TL_APMU_VREF_P40 As Long = 132
Public Const TL_APMU_VREF_M2 As Long = 18
Public Const TL_APMU_VREF_M5 As Long = 34
Public Const TL_APMU_VREF_M10 As Long = 66
Public Const TL_APMU_VREF_M40 As Long = 130
Public Const TL_APMU_VREF_MECCA As Long = 8
Public Const TL_APMU_VREFA_P10 As Long = 2
Public Const TL_APMU_VREFA_M10 As Long = 4

Public Const TL_APMU_CALSRC_NONE As Long = &H0
Public Const TL_APMU_CALSRC_VREF As Long = &H1
Public Const TL_APMU_CALSRC_RREF As Long = &H2
Public Const TL_APMU_CALSRC_RREF_L As Long = &H4
Public Const TL_APMU_CALSRC_MECCA As Long = &H8

Public Const TL_APMU_CAL_SLOT_ONLY As Long = &HFF000000
Public Const TL_APMU_CAL_CHANMAP_ONLY As Long = &HFFFFFFFE
Public Const TL_APMU_CAL_ALL_CHANNEL As Long = &HFFFFFFFF

Public Const TL_APMU_CAL_NORMAL As Long = 0
Public Const TL_APMU_CAL_NOMINAL As Long = 1
Public Const TL_APMU_CAL_FORCEREAD As Long = 2

