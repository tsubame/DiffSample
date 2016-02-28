VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DCTestScenario_IE 
   Caption         =   "InstanceEditor"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12150
   OleObjectBlob   =   "DCTestScenario_IE.frx":0000
   StartUpPosition =   1  'ÉIÅ[ÉiÅ[ ÉtÉHÅ[ÉÄÇÃíÜâõ
End
Attribute VB_Name = "DCTestScenario_IE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





'äTóv:
'   DC Test ScenarioópÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^
'
'ñ⁄ìI:
'   Eee-JOB DCTestScenarioÉeÉXÉgÉCÉìÉXÉ^ÉìÉXÉtÉåÅ[ÉÄÉèÅ[ÉNópÇ…êÍópÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ÇçÏê¨
'   IG-XL Ver.3.40.17ÇÃInstanceEditor_IEÇÉxÅ[ÉXÇ…ÅAåªíiäKÇ≈ïsóvÇ»ÉpÉâÉÅÅ[É^Ç…ä÷ÇµÇƒÅAÉtÉHÅ[ÉÄè„ÇÃ
'   ÉRÉìÉgÉçÅ[ÉãÇÕîÒï\é¶Ç…ÇµÇƒÇªÇÍÇ…çáÇÌÇπÇƒÉTÉCÉYÇÃïœçXÇçsÇ§èàóùÇí«â¡ÇµÉJÉXÉ^É}ÉCÉYÇçsÇ¡ÇΩ
'
'   è´óàìIÇ…ÉeÉìÉvÉåÅ[Égè„Ç≈É^ÉCÉ~ÉìÉOÅAÉpÉ^Å[ÉìÅAÉsÉìÇÃê›íËÇçsÇ§â¬î\ê´ÇécÇ∑ÇΩÇﬂÇ…ÅA
'   ÉRÉìÉgÉçÅ[ÉãÅATemplateArgÉIÉuÉWÉFÉNÉgÇÃê›íËÇÕäÓñ{ìIÇ…ïœçXÇµÇƒÇ¢Ç»Ç¢
'
'   ÇΩÇæÇµÅAIG-XL Ver.3.40.10ê¢ë„Ç∆å›ä∑ê´ÇéùÇΩÇπÇÈÇΩÇﬂÇ…à»â∫ÇÃïœçXÇçsÇ¡ÇƒÇ¢ÇÈ
'
'   á@Ver.3.40.10ê¢ë„ÇÕÅuCommentsÅvÉeÉLÉXÉgÉ{ÉbÉNÉXÉRÉìÉgÉçÅ[ÉãÇÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂÅA
'   Å@tl_tm_ParFuncCommentsTextBoxÉIÉuÉWÉFÉNÉgÇÃê›íËÇïœçX
'   áAVer.3.40.10ê¢ë„ÇÕOptionalArgumentsÉtÉåÅ[ÉÄÇÃPatternÉyÅ[ÉWÇÃ
'   Å@ÅuPattern Groups & SetsÅvÉRÉìÉgÉçÅ[ÉãÉ{É^ÉìÇÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂÅA
'   Å@tl_tm_ParPrecondPatNamesÉIÉuÉWÉFÉNÉgãyÇ—tl_tm_ParHoldStatePatNameÉIÉuÉWÉFÉNÉgÇ÷ÇÃ
'   Å@ÉRÉìÉgÉçÅ[ÉãÉ{É^ÉìÇÃí«â¡ï˚ñ@ÇïœçX
'   áBVer.3.40.10ê¢ë„ÇÕÅuOptional ArgumentsÅvÉtÉåÅ[ÉÄÇÃÅuDPSÅvÉyÅ[ÉWÇÃ
'   Å@ÅuPre Cond Pat ClampÅvÉeÉLÉXÉgÉ{ÉbÉNÉXÅiPowerSupply_TÇ≈égópÅjÇÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂ
'   Å@tl_tm_ParDpsprecondpatClampÉIÉuÉWÉFÉNÉgê›íËÇÉRÉÅÉìÉgÉAÉEÉg
'
'   [íçà”éñçÄ]
'   Å@ÅuCommentsÅvÉeÉLÉXÉgÉ{ÉbÉNÉXÉRÉìÉgÉçÅ[ÉãÇVer.3.40.10ê¢ë„Ç…Ç‡ëŒâûÇ≥ÇπÇÈÇΩÇﬂÇ…ÅA
'   Å@ë„ÇÌÇËÇ…ÉÜÅ[ÉUÅ[ÉJÉXÉ^ÉÄópTemplateArgÉIÉuÉWÉFÉNÉgïœêî[tl_tm_ParUserOpt10]Çâ°éÊÇËÇµÇƒÇ¢ÇÈ
'   Å@ÇªÇÃÇΩÇﬂÇ±ÇÃÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^Ç≈ÇÕ10î‘ñ⁄ÇÃÉÜÅ[ÉUÅ[ÉIÉvÉVÉáÉìïœêîÇÕégópèoóàÇ»Ç¢Ç±Ç∆Ç…íçà”
'
'çÏê¨é“:
'   0145206097
'





' (c) Teradyne, Inc, 1997, 1998, 1999
'     All Rights Reserved
' Inclusion of a copyright notice does not imply that this software has been
' published.
' This software is the trade secret information of Teradyne, Inc.
' Use of this software is only in accordance with the terms of a license
' agreement from Teradyne, Inc.
'
' Revision History:
' Date        Description
' 01/16/07      Prabha - To add a new module HDVIS power supply.
' 04/03/06     Ganesh Pandiyan K Fix tersw00072341 - Added an Edit box to enter comments in the Template GUI.
' 04/03/06    Vidhya.R    Fix for the defect TERSW00084751-Instances sheet ñ resize the instance dialog boxes so all information is
'             visible.
'01/17/06     Mari Selvi.P Fix TERW00074157 - Added Button under patterns tab.
' 08/08/05    Boopathi P  Fix for TERSW00059979 - Added a text box to PowerSupply_T template
'             Allow the user to set a precondition pattern current Clamp value
' 09/27/99    Release 3.30 Development
' 03/22/02    CQ7536 - Leave EdgeSet in Black when selecting an empty string.

Option Explicit

'***************************************************
'NEXT 1 IS USED TO DENOTE TEMPLATE PARAMETERS
'***************************************************
Private Sub ParameterSetup()
    Dim intX As Integer
    Dim temp As String

    '------------------------------------------------------------------------------------
    'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
    'IG-XL Ver3.40.10ê¢ë„Ç≈ÇÕtl_tm_ParFuncCommentsTextBoxÇÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂÅA
    'ë„ÇÌÇËÇ…ÉÜÅ[ÉUÅ[ÉJÉXÉ^ÉÄópTemplateArgÉIÉuÉWÉFÉNÉgïœêî[tl_tm_ParUserOpt10]Çîqéÿ
    Set tl_tm_ParUserOpt10.DataBox = CommentsTextBox
    'Fix for the defect tersw00072341 - Added an edit box for entering comments.
'    With tl_tm_ParFuncCommentsTextBox
'        Set .DataBox = CommentsTextBox
'    End With
    '------------------------------------------------------------------------------------

    'Provide the references to the Labels and Databoxes for many commonly used Parameters
    If Me.OptPages.LevTimPage.Visible Then
        Call tl_tm_SetupCatSelObjects(DcCatCombo, DcCatLabel, AcCatCombo, AcCatLabel, _
            DcSelCombo, DcSelLabel, AcSelCombo, AcSelLabel)
        Call tl_tm_SetupTimLevObjects(TimeSetCombo, TimeSetLabel, EdgeSetCombo, EdgeSetLabel, _
            LevelsCombo, LevelsLabel)
        Call tl_tm_SetupOverlayObjects(OverlayCombo, OverlayLabel)
    End If
    If Me.OptPages.InterposePage.Visible Then
        If Me.SerializeMeasFBox.Visible Then
            '----------------------------------------------------------------------------
            'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
            'ÉGÉâÅ[âÒîÇÃÇΩÇﬂèàóùïœçX
            'Ç±ÇÃÉGÉfÉBÉ^Ç≈ÇÕégópÇµÇ»Ç¢ÉpÉâÉÅÅ[É^ÉIÉuÉWÉFÉNÉgÇ…ÉfÅ[É^É{ÉbÉNÉXÉRÉìÉgÉçÅ[ÉãÇí«â¡ÇµÇƒÇµÇ‹Ç§Ç∆
            'ëºÇÃÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ÇãNìÆÇ≥ÇπÇΩÇ∆Ç´Ç…ÉGÉâÅ[Ç…Ç»Ç¡ÇƒÇµÇ‹Ç§ÇΩÇﬂÅA[Nothing]Çó\Çﬂê›íËÇµÇƒÇ®Ç≠
'            Call tl_tm_SetupInterposeObjects(StartOfBodyBox, StartOfBodyFLabel, PrePatFBox, PrePatFLabel, _
'                PreTestBox, PreTestFLabel, PostTestBox, PostTestFLabel, _
'                PostPatFBox, PostPatFLabel, EndOfBodyBox, EndOfBodyFLabel, _
'                SerializeMeasFBox, SerializeMeasFLabel)
'            Call tl_tm_SetupInterposeInputObjects(StartOfBodyInputBox, PrePatInputBox, _
'                PreTestInputBox, PostTestInputBox, _
'                PostPatInputBox, EndOfBodyInputBox, _
'                SerializeMeasFInputBox)
            Call tl_tm_SetupInterposeObjects(StartOfBodyBox, StartOfBodyFLabel, Nothing, Nothing, _
                Nothing, Nothing, Nothing, Nothing, _
                Nothing, Nothing, EndOfBodyBox, EndOfBodyFLabel, _
                Nothing, Nothing)
            Call tl_tm_SetupInterposeInputObjects(StartOfBodyInputBox, Nothing, _
                Nothing, Nothing, _
                Nothing, EndOfBodyInputBox, _
                Nothing)
            '----------------------------------------------------------------------------
        Else
            '----------------------------------------------------------------------------
            'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
            'ÉGÉâÅ[âÒîÇÃÇΩÇﬂèàóùïœçX
            'Ç±ÇÃÉGÉfÉBÉ^Ç≈ÇÕégópÇµÇ»Ç¢ÉpÉâÉÅÅ[É^ÉIÉuÉWÉFÉNÉgÇ…ÉfÅ[É^É{ÉbÉNÉXÉRÉìÉgÉçÅ[ÉãÇí«â¡ÇµÇƒÇµÇ‹Ç§Ç∆
            'ëºÇÃÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ÇãNìÆÇ≥ÇπÇΩÇ∆Ç´Ç…ÉGÉâÅ[Ç…Ç»Ç¡ÇƒÇµÇ‹Ç§ÇΩÇﬂÅA[Nothing]Çó\Çﬂê›íËÇµÇƒÇ®Ç≠
'            Call tl_tm_SetupInterposeObjects(StartOfBodyBox, StartOfBodyFLabel, PrePatFBox, PrePatFLabel, _
'                PreTestBox, PreTestFLabel, PostTestBox, PostTestFLabel, _
'                PostPatFBox, PostPatFLabel, EndOfBodyBox, EndOfBodyFLabel)
'            Call tl_tm_SetupInterposeInputObjects(StartOfBodyInputBox, PrePatInputBox, _
'                PreTestInputBox, PostTestInputBox, _
'                PostPatInputBox, EndOfBodyInputBox)
            Call tl_tm_SetupInterposeObjects(StartOfBodyBox, StartOfBodyFLabel, Nothing, Nothing, _
                Nothing, Nothing, Nothing, Nothing, _
                Nothing, Nothing, EndOfBodyBox, EndOfBodyFLabel)
            Call tl_tm_SetupInterposeInputObjects(StartOfBodyInputBox, Nothing, _
                Nothing, Nothing, _
                Nothing, EndOfBodyInputBox)
            '----------------------------------------------------------------------------
        End If
    End If
    If Me.OptPages.PatFlagFuncPage.Visible Then
        'WaitFlags,
        With tl_tm_ParPatFlags
            Set .LabelBox = PatWaitFlagsLabel
        End With
        'FlagWaitTimeout,
        With tl_tm_ParPatFlagWaitTime
            Set .DataBox = PatFlagWaitTimeBox
            Set .LabelBox = PatFlagWaitTimeoutLabel
            Set .EvalBox = PatFlagWaitEvalLabel
        End With
        'PatFlagF,
        With tl_tm_ParPatFlagF
            Set .DataBox = PatFunctionBox
            Set .LabelBox = PatFunctionLabel
        End With
        'PatFlagFInput,
        With tl_tm_ParPatFlagFInput
            Set .DataBox = PatFunctionInput
            Set .LabelBox = Nothing
        End With
        'MatchAllSites,
        With tl_tm_ParPatMatchAllSites
            Set .DataBox = MatchAllSitesCombo
            Set .LabelBox = MatchAllSitesLabel
        End With
    End If
    If Me.OptPages.PinPage.Visible Then
        Call tl_tm_SetupConditioningPinlistsObjects(StartInitLOBox, StartLOLabel, StartInitHIBox, StartHILabel, _
            StartInitZBox, StartZLabel, FloatPinsBox, FloatLabel, Util1PinsBox, Util1Label, Util0PinsBox, Util0Label, _
            Type3Button, FloatEditButton, UtilEditButton)
    End If
    If Me.OptPages.PatPage.Visible Then

        '--------------------------------------------------------------------------------
        'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
        'IG-XL Ver3.40.10ê¢ë„Ç…ëŒâûÇ∑ÇÈÇΩÇﬂÅA
        'tl_tm_ParPrecondPatNamesÉIÉuÉWÉFÉNÉgÇ÷ÇÃPcpGrpAndSetButtonÉ{É^ÉìÉRÉìÉgÉçÅ[ÉãÇÃí«â¡Ç
        'tl_tm_SetupPreCondPatObjectsÅiValSupportÉÇÉWÉÖÅ[ÉãÅjä÷êîÇ©ÇÁêÿÇËèoÇµ
        Call tl_tm_SetupPreCondPatObjects(PreCondPatCombo, PreCondPatLabel, PcpStopLabelBox, PcpStopLabelLabel, _
            PcpStartLabelBox, PcpStartLabelLabel, PcpCheckPgCombo, PcpCheckPGLabel, PcpFinderButton) ' , PcpGrpAndSetButton)
        Set tl_tm_ParPrecondPatNames.SelectorButton = PcpGrpAndSetButton
        '--------------------------------------------------------------------------------

        '--------------------------------------------------------------------------------
        'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
        'IG-XL Ver3.40.10ê¢ë„Ç…ëŒâûÇ∑ÇÈÇΩÇﬂÅA
        'tl_tm_ParHoldStatePatNameÉIÉuÉWÉFÉNÉgÇ÷ÇÃHspGrpAndSetButtonÉ{É^ÉìÉRÉìÉgÉçÅ[ÉãÇÃí«â¡Ç
        'tl_tm_SetupHoldStatePatObjectsÅiValSupportÉÇÉWÉÖÅ[ÉãÅjä÷êîÇ©ÇÁêÿÇËèoÇµ
        Call tl_tm_SetupHoldStatePatObjects(HoldStatePatCombo, HoldStatePatLabel, HspStopLabelBox, HspStopLabelLabel, _
            HspStartLabelBox, HspStartLabelLabel, HspCheckPgCombo, HspCheckPGLabel, HspFinderButton, _
            HspResumeCombo, HspResumeLabel, WaitFlagsLabel, FlagWaitTimeBox, FlagWaitTimeoutLabel, FlagWaitEvalLabel) ', HspGrpAndSetButton)
        Set tl_tm_ParHoldStatePatName.SelectorButton = HspGrpAndSetButton
        '--------------------------------------------------------------------------------

    End If
    If Me.ReqPages.PatPage.Visible Then
        'Provide the references to the Labels and Databoxes for Parameters unique to this template
        'Patterns,
        With tl_tm_ParFuncPatternName
            Set .DataBox = FuncPatternTextBox
            Set .LabelBox = FuncPatternsLabel
            Set .FileFindButton = FuncFileFinderButton
            Set .SelectorButton = FuncPatListEditButton
        End With
        'SetPassFail,
        With tl_tm_ParFuncSetPassFail
            Set .DataBox = FuncSetPassFailCombo
            Set .LabelBox = FuncSetPassFailLabel
        End With
        'RelayMode,
        With tl_tm_ParFuncRelayMode
            Set .DataBox = FuncRelayModeCombo
            Set .LabelBox = FuncRelayModeLabel
        End With
        'Threading,
        With tl_tm_ParFuncThreading
            Set .DataBox = ThreadingCombo
            Set .LabelBox = ThreadingLabel
        End With
    End If
    If Me.ReqPages.DpsPage.Visible Then
        'Provide the references to the Labels and Databoxes for Parameters unique to this template
        '(DPS items whether required or not are handled here because both Req and Opt pages are
        '   set if the other is.
        'TestControl,
        With tl_tm_ParDpsTestControl
            Set .DataBox = DpsTestControlCombo
            Set .LabelBox = DpsTestControlLabel
        End With
        'Irange,
        With tl_tm_ParDpsIRange
            Set .DataBox = DpsIRangeCombo
            Set .LabelBox = DpsIRangeLabel
        End With
        'Clamp,
        With tl_tm_ParDpsClamp
            Set .DataBox = DpsClampBox
            Set .EvalBox = DpsClampEvalLabel
            Set .LabelBox = DpsClampLabel
        End With

        '--------------------------------------------------------------------------------
        'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
        'IG-XL Ver3.40.10ê¢ë„Ç≈ÇÕPowerSupply_TÇ≈tl_tm_ParDpsprecondpatClampÇ
        'ÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂÉRÉÅÉìÉgÉAÉEÉg
        'Fix for TERSW00059979 - Added a text box to PowerSupply_T template
        'Pre cond Pat Clamp
'        With tl_tm_ParDpsprecondpatClamp
'            Set .DataBox = DpsPrecondpatClampBox
'            Set .EvalBox = DpsPrecondpatClampEvalLabel
'            Set .LabelBox = DpsPrecondpatClampLabel
'        End With
        '--------------------------------------------------------------------------------

        'SamplingTime,
        With tl_tm_ParDpsSamplingTime
            Set .DataBox = DpsSamplingTimeBox
            Set .EvalBox = DpsSampTimeEvalLabel
            Set .LabelBox = DpsSamplingTimeLabel
        End With
        'Samples,
        With tl_tm_ParDpsSamples
            Set .DataBox = DpsSamplesBox
            Set .EvalBox = DpsSampEvalLabel
            Set .LabelBox = DpsSamplesLabel
        End With
        'SettlingTime,
        With tl_tm_ParDpsSettlingTime
            Set .DataBox = DpsSettlingTimeBox
            Set .EvalBox = DpsSetTimeEvalLabel
            Set .LabelBox = DpsSettlingTimeLabel
        End With
        'HiLoLimValid,
        'HiLimit,
        With tl_tm_ParDpsHiLimSpec
            Set .DataBox = DpsHiLimSpecCombo
            Set .EvalBox = DpsHiLimEvalLabel
            Set .LabelBox = DpsHiLimSpecLabel
            Set .EquationButton = DpsHiLimEquationButton
        End With
        'LoLimit,
        With tl_tm_ParDpsLoLimSpec
            Set .DataBox = DpsLoLimSpecCombo
            Set .EvalBox = DpsLoLimEvalLabel
            Set .LabelBox = DpsLoLimSpecLabel
            Set .EquationButton = DpsLoLimEquationButton
        End With
        'MainForceCond,
        With tl_tm_ParDpsMainForceCond
            Set .DataBox = DpsMainForceCondCombo
            Set .EvalBox = DpsMainForceCondEvalLabel
            Set .LabelBox = DpsMainForceCondLabel
            Set .EquationButton = DpsMainForceCondEquationButton
        End With
        'AltForceCond,
        With tl_tm_ParDpsAltForceCond
            Set .DataBox = DpsAltForceCondCombo
            Set .EvalBox = DpsAltForceCondEvalLabel
            Set .LabelBox = DpsAltForceCondLabel
            Set .EquationButton = DpsAltForceCondEquationButton
        End With
        'PowerPin,
        With tl_tm_ParDpsPowerPin
            Set .DataBox = DpsPowerPinListBox
            Set .LabelBox = DpsPowerPinLabel
            Set .SelectorButton = DpsPowerPinListEditButton
        End With
        'MainOrAlt
        With tl_tm_ParDpsMAINORALT
            Set .DataBox = DpsMainOrAltCombo
            Set .LabelBox = DpsMainOrAltLabel
        End With
        'RelayMode,
        With tl_tm_ParDpsRelayMode
            Set .DataBox = DpsRelayModeCombo
            Set .LabelBox = DpsRelayModeLabel
        End With
        'Vrange,
        'no parameter by this name appears on the IE form, this is used solely for validation purposes
        With tl_tm_ParDpsVrange
            Set .LabelBox = Nothing
            Set .DataBox = Nothing
        End With
        
        With tl_tm_ParSerializeMeas
           Set .CheckBox = SerializeMeasCheckBox
        End With
           
    End If
    If Me.ReqPages.PmuPage.Visible Then
        'Provide the references to the Labels and Databoxes for Parameters unique to this IE page
        If tl_tm_InstanceEditor.BpmuPages Then
            'Pinlist,
            With tl_tm_ParBpmuPinlist
                Set .DataBox = PmuPinListBox
                Set .LabelBox = PmuPinListLabel
                Set .SelectorButton = PmuPinListEditButton
            End With
            'HiLoLimValid,
            'HiLimit,
            With tl_tm_ParBpmuHiLimSpec
                Set .DataBox = PmuHiLimSpecCombo
                Set .EvalBox = PmuHiLimEvalLabel
                Set .LabelBox = PmuHiLimSpecLabel
                Set .EquationButton = PmuHiLimEquationButton
            End With
            'LoLimit,
            With tl_tm_ParBpmuLoLimSpec
                Set .DataBox = PmuLoLimSpecCombo
                Set .EvalBox = PmuLoLimEvalLabel
                Set .LabelBox = PmuLoLimSpecLabel
                Set .EquationButton = PmuLoLimEquationButton
            End With
            'ForceCond1,
            With tl_tm_ParBpmuForceCond1
                Set .DataBox = PmuForceCond1Combo
                Set .EvalBox = PmuForceCond1EvalLabel
                Set .LabelBox = PmuForceCond1Label
                Set .EquationButton = PmuForceCond1EquationButton
            End With
        Else
            'Pinlist,
            With tl_tm_ParPpmuPinlist
                Set .DataBox = PmuPinListBox
                Set .LabelBox = PmuPinListLabel
                Set .SelectorButton = PmuPinListEditButton
            End With
            'HiLoLimValid,
            'HiLimit,
            With tl_tm_ParPpmuHiLimSpec
                Set .DataBox = PmuHiLimSpecCombo
                Set .EvalBox = PmuHiLimEvalLabel
                Set .LabelBox = PmuHiLimSpecLabel
                Set .EquationButton = PmuHiLimEquationButton
            End With
            'LoLimit,
            With tl_tm_ParPpmuLoLimSpec
                Set .DataBox = PmuLoLimSpecCombo
                Set .EvalBox = PmuLoLimEvalLabel
                Set .LabelBox = PmuLoLimSpecLabel
                Set .EquationButton = PmuLoLimEquationButton
            End With
            'ForceCond1,
            With tl_tm_ParPpmuForceCond1
                Set .DataBox = PmuForceCond1Combo
                Set .EvalBox = PmuForceCond1EvalLabel
                Set .LabelBox = PmuForceCond1Label
                Set .EquationButton = PmuForceCond1EquationButton
            End With
        End If
    End If
    If Me.OptPages.BpmuPage.Visible Then
        'Provide the references to the Labels and Databoxes for Parameters unique to this IE page
        'MeasureMode,
        With tl_tm_ParBpmuMeasureMode
            Set .DataBox = BpmuMeasureModeCombo
            Set .LabelBox = BpmuMeasureModeLabel
        End With
        'Irange,
        With tl_tm_ParBpmuIRange
            Set .DataBox = BpmuIRangeCombo
            Set .LabelBox = BpmuIRangeLabel
        End With
        'Clamp,
        With tl_tm_ParBpmuClamp
            Set .DataBox = BpmuClampBox
            Set .EvalBox = BpmuClampEvalLabel
            Set .LabelBox = BpmuClampLabel
        End With
        'Vrange,
        With tl_tm_ParBpmuVrange
            Set .DataBox = BpmuVRangeCombo
            Set .LabelBox = BpmuVRangeLabel
        End With
        'SamplingTime,
        With tl_tm_ParBpmuSamplingTime
            Set .DataBox = BpmuSamplingTimeBox
            Set .EvalBox = BpmuSampTimeEvalLabel
            Set .LabelBox = BpmuSamplingTimeLabel
        End With
        'Samples,
        With tl_tm_ParBpmuSamples
            Set .DataBox = BpmuSamplesBox
            Set .EvalBox = BpmuSampEvalLabel
            Set .LabelBox = BpmuSamplesLabel
        End With
        'SettlingTime,
        With tl_tm_ParBpmuSettlingTime
            Set .DataBox = BpmuSettlingTimeBox
            Set .EvalBox = BpmuSetTimeEvalLabel
            Set .LabelBox = BpmuSettlingTimeLabel
        End With
        'ForceCond2,
        With tl_tm_ParBpmuForceCond2
            Set .DataBox = BpmuForceCond2Combo
            Set .EvalBox = BpmuForceCond2EvalLabel
            Set .LabelBox = BpmuForceCond2Label
            Set .EquationButton = BpmuForceCond2EquationButton
        End With
        'GangPinsTested
        With tl_tm_ParBpmuGang
            Set .DataBox = BpmuGangCombo
            Set .LabelBox = BpmuGangPinsTestedLabel
        End With
        'RelayMode,
        With tl_tm_ParBpmuRelayMode
            Set .DataBox = BpmuRelayModeCombo
            Set .LabelBox = BpmuRelayModeLabel
        End With
    End If
    If Me.OptPages.PpmuPage.Visible Then
        'Provide the references to the Labels and Databoxes for Parameters unique to this IE page
        'MeasureMode,
        With tl_tm_ParPpmuMeasureMode
            Set .DataBox = PpmuMeasureModeCombo
            Set .LabelBox = PpmuMeasureModeLabel
        End With
        'Irange,
        With tl_tm_ParPpmuIRange
            Set .DataBox = PpmuIRangeCombo
            Set .LabelBox = PpmuIRangeLabel
        End With
        'HClamp,
        With tl_tm_ParPpmuHClamp
            Set .DataBox = PpmuHClampBox
            Set .EvalBox = PpmuHClampEvalLabel
            Set .LabelBox = PpmuHClampLabel
        End With
        'LClamp,
        With tl_tm_ParPpmuLClamp
            Set .DataBox = PpmuLClampBox
            Set .EvalBox = PpmuLClampEvalLabel
            Set .LabelBox = PpmuLClampLabel
        End With
        'SamplingTime,
        With tl_tm_ParPpmuSamplingTime
            Set .DataBox = PpmuSamplingTimeBox
            Set .EvalBox = PpmuSampTimeEvalLabel
            Set .LabelBox = PpmuSamplingTimeLabel
        End With
        'Samples,
        With tl_tm_ParPpmuSamples
            Set .DataBox = PpmuSamplesBox
            Set .EvalBox = PpmuSampEvalLabel
            Set .LabelBox = PpmuSamplesLabel
        End With
        'SettlingTime,
        With tl_tm_ParPpmuSettlingTime
            Set .DataBox = PpmuSettlingTimeBox
            Set .EvalBox = PpmuSetTimeEvalLabel
            Set .LabelBox = PpmuSettlingTimeLabel
        End With
        'ForceCond2,
        With tl_tm_ParPpmuForceCond2
            Set .DataBox = PpmuForceCond2Combo
            Set .EvalBox = PpmuForceCond2EvalLabel
            Set .LabelBox = PpmuForceCond2Label
            Set .EquationButton = PpmuForceCond2EquationButton
        End With
        'FLoad,
        With tl_tm_ParPpmuFLoad
            Set .DataBox = PpmuFloadCombo
            Set .LabelBox = PpmuFloadLabel
        End With
        'RelayMode,
        With tl_tm_ParPpmuRelayMode
            Set .DataBox = PpmuRelayModeCombo
            Set .LabelBox = PpmuRelayModeLabel
        End With
        'Vrange,
        'no parameter by this name appears on the IE form, this is used solely for validation purposes
        '''''hey hey  tl_tm_ParPpmuVrange
        With tl_tm_ParPpmuVrange
            .ParameterStr = TL_C_EMPTYSTR
            Set .LabelBox = Nothing
            Set .DataBox = Nothing
        End With
    End If
    
    With Me
        If .ReqPages.UserPage1.Visible Then
            'set the name of the page
            If tl_tm_InstanceEditor.UserReqName <> TL_C_EMPTYSTR Then
                .ReqPages.UserPage1.Caption = tl_tm_InstanceEditor.UserReqName
            End If
        End If
        If .ReqPages.UserPage2.Visible Then
            'set the name of the page
            If (tl_tm_InstanceEditor.UserReqName <> TL_C_EMPTYSTR) And (tl_tm_InstanceEditor.UserReqName2 = TL_C_EMPTYSTR) Then
                .ReqPages.UserPage2.Caption = tl_tm_InstanceEditor.UserReqName
            End If
            If tl_tm_InstanceEditor.UserReqName2 <> TL_C_EMPTYSTR Then
                .ReqPages.UserPage2.Caption = tl_tm_InstanceEditor.UserReqName2
            End If
        End If
        If .OptPages.UserPage1.Visible Then
            'set the name of the page
            If tl_tm_InstanceEditor.UserOptName <> TL_C_EMPTYSTR Then
                .OptPages.UserPage1.Caption = tl_tm_InstanceEditor.UserOptName
            End If
        End If
        If .OptPages.UserPage2.Visible Then
            'set the name of the page
            If (tl_tm_InstanceEditor.UserOptName <> TL_C_EMPTYSTR) And (tl_tm_InstanceEditor.UserOptName2 = TL_C_EMPTYSTR) Then
                .OptPages.UserPage2.Caption = tl_tm_InstanceEditor.UserOptName
            End If
            If tl_tm_InstanceEditor.UserOptName2 <> TL_C_EMPTYSTR Then
                .OptPages.UserPage2.Caption = tl_tm_InstanceEditor.UserOptName2
            End If
        End If
        
        'set all objects invisible
        Call tl_fs_SetPageObjectInvisible(.ReqPages.UserPage1)
        Call tl_fs_SetPageObjectInvisible(.ReqPages.UserPage2)
        Call tl_fs_SetPageObjectInvisible(.OptPages.UserPage1)
        Call tl_fs_SetPageObjectInvisible(.OptPages.UserPage2)
    End With
    
    'now conditionally set them visible
    'set up the proper objects and sizes and references
    If tl_tm_InstanceEditor.UserReqArg1.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq1, tl_tm_InstanceEditor.UserReqArg1, _
        UserReqLabel1, UserReqComboBox1, UserReqTextBox1, _
        UserReqButton1, UserReqEvalLabel1, UserReqCheckBox1)
    End If
    If tl_tm_InstanceEditor.UserReqArg2.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq2, tl_tm_InstanceEditor.UserReqArg2, _
        UserReqLabel2, UserReqComboBox2, UserReqTextBox2, _
        UserReqButton2, UserReqEvalLabel2, UserReqCheckBox2)
    End If
    If tl_tm_InstanceEditor.UserReqArg3.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq3, tl_tm_InstanceEditor.UserReqArg3, _
        UserReqLabel3, UserReqComboBox3, UserReqTextBox3, _
        UserReqButton3, UserReqEvalLabel3, UserReqCheckBox3)
    End If
    If tl_tm_InstanceEditor.UserReqArg4.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq4, tl_tm_InstanceEditor.UserReqArg4, _
        UserReqLabel4, UserReqComboBox4, UserReqTextBox4, _
        UserReqButton4, UserReqEvalLabel4, UserReqCheckBox4)
    End If
    If tl_tm_InstanceEditor.UserReqArg5.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq5, tl_tm_InstanceEditor.UserReqArg5, _
        UserReqLabel5, UserReqComboBox5, UserReqTextBox5, _
        UserReqButton5, UserReqEvalLabel5, UserReqCheckBox5)
    End If
    If tl_tm_InstanceEditor.UserReqArg6.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq6, tl_tm_InstanceEditor.UserReqArg6, _
        UserReqLabel6, UserReqComboBox6, UserReqTextBox6, _
        UserReqButton6, UserReqEvalLabel6, UserReqCheckBox6)
    End If
    If tl_tm_InstanceEditor.UserReqArg7.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq7, tl_tm_InstanceEditor.UserReqArg7, _
        UserReqLabel7, UserReqComboBox7, UserReqTextBox7, _
        UserReqButton7, UserReqEvalLabel7, UserReqCheckBox7)
    End If
    If tl_tm_InstanceEditor.UserReqArg8.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq8, tl_tm_InstanceEditor.UserReqArg8, _
        UserReqLabel8, UserReqComboBox8, UserReqTextBox8, _
        UserReqButton8, UserReqEvalLabel8, UserReqCheckBox8)
    End If
    If tl_tm_InstanceEditor.UserReqArg9.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq9, tl_tm_InstanceEditor.UserReqArg9, _
        UserReqLabel9, UserReqComboBox9, UserReqTextBox9, _
        UserReqButton9, UserReqEvalLabel9, UserReqCheckBox9)
    End If
    If tl_tm_InstanceEditor.UserReqArg10.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserReq10, tl_tm_InstanceEditor.UserReqArg10, _
        UserReqLabel10, UserReqComboBox10, UserReqTextBox10, _
        UserReqButton10, UserReqEvalLabel10, UserReqCheckBox10)
    End If
    
    If tl_tm_InstanceEditor.UserOptArg1.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt1, tl_tm_InstanceEditor.UserOptArg1, _
        UserOptLabel1, UserOptComboBox1, UserOptTextBox1, _
        UserOptButton1, UserOptEvalLabel1, UserOptCheckBox1)
    End If
    If tl_tm_InstanceEditor.UserOptArg2.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt2, tl_tm_InstanceEditor.UserOptArg2, _
        UserOptLabel2, UserOptComboBox2, UserOptTextBox2, _
        UserOptButton2, UserOptEvalLabel2, UserOptCheckBox2)
    End If
    If tl_tm_InstanceEditor.UserOptArg3.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt3, tl_tm_InstanceEditor.UserOptArg3, _
        UserOptLabel3, UserOptComboBox3, UserOptTextBox3, _
        UseroptButton3, UserOptEvalLabel3, UserOptCheckBox3)
    End If
    If tl_tm_InstanceEditor.UserOptArg4.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt4, tl_tm_InstanceEditor.UserOptArg4, _
        UserOptLabel4, UserOptComboBox4, UserOptTextBox4, _
        UserOptButton4, UserOptEvalLabel4, UserOptCheckBox4)
    End If
    If tl_tm_InstanceEditor.UserOptArg5.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt5, tl_tm_InstanceEditor.UserOptArg5, _
        UserOptLabel5, UserOptComboBox5, UserOptTextBox5, _
        UserOptButton5, UserOptEvalLabel5, UserOptCheckBox5)
    End If
    If tl_tm_InstanceEditor.UserOptArg6.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt6, tl_tm_InstanceEditor.UserOptArg6, _
        UserOptLabel6, UserOptComboBox6, UserOptTextBox6, _
        UserOptButton6, UserOptEvalLabel6, UserOptCheckBox6)
    End If
    If tl_tm_InstanceEditor.UserOptArg7.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt7, tl_tm_InstanceEditor.UserOptArg7, _
        UserOptLabel7, UserOptComboBox7, UserOptTextBox7, _
        UserOptButton7, UserOptEvalLabel7, UserOptCheckBox7)
    End If
    If tl_tm_InstanceEditor.UserOptArg8.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt8, tl_tm_InstanceEditor.UserOptArg8, _
        UserOptLabel8, UserOptComboBox8, UserOptTextBox8, _
        UserOptButton8, UserOptEvalLabel8, UserOptCheckBox8)
    End If
    If tl_tm_InstanceEditor.UserOptArg9.enabled Then
        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt9, tl_tm_InstanceEditor.UserOptArg9, _
        UserOptLabel9, UserOptComboBox9, UserOptTextBox9, _
        UserOptButton9, UserOptEvalLabel9, UserOptCheckBox9)
    End If

    '------------------------------------------------------------------------------------
    'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
    'IG-XL Ver3.40.10ê¢ë„Ç≈ÇÕtl_tm_ParFuncCommentsTextBoxÇÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂÅA
    'ë„ÇÌÇËÇ…ÉÜÅ[ÉUÅ[ÉJÉXÉ^ÉÄópTemplateArgÉIÉuÉWÉFÉNÉgïœêî[tl_tm_ParUserOpt10]Çîqéÿ
    'ñ{óàÇÃtl_tm_ParUserOpt10ïœêîÇ÷ÇÃÉtÉHÅ[ÉÄÉRÉìÉgÉçÅ[ÉãÇÃê›íËÇÕÉRÉÅÉìÉgÉAÉEÉg
'    If tl_tm_InstanceEditor.UserOptArg10.Enabled Then
'        Call tl_fs_SetupCustomDisplay(tl_tm_ParUserOpt10, tl_tm_InstanceEditor.UserOptArg10, _
'        UserOptLabel10, UserOptComboBox10, UserOptTextBox10, _
'        UserOptButton10, UserOptEvalLabel10, UserOptCheckBox10)
'    End If
    '------------------------------------------------------------------------------------

    'call the SetupParameters routine for the Teradyne template, or Custom template
    temp = tl_dt_IEGetTemplateName
    intX = InStr(temp, "!")
    If intX <> 0 Then
        temp = temp & ".SetupParameters"
        On Error GoTo RunErr
        'run the proper SetupParameters routine!
        Run temp
        On Error GoTo 0
    Else
        'denote an error
        Call tl_ErrorLogMessage("InstanceEditor: tl_dt_IEGetTemplateName " & TL_C_ERRORSTR & " : " & temp)
        Call tl_ErrorReport
    End If
    
    Exit Sub
RunErr:
    On Error GoTo 0
    'denote error
    Call tl_ErrorLogMessage("InstanceEditor: SetupParameters " & TL_C_ERRORSTR)
    Call tl_ErrorReport
End Sub



'***************************************************
'NEXT 6 HANDLE BUTTON CLICKS
'***************************************************
Private Sub ApplyButton_Click()
    If tl_tm_FormCtrl.ContextChanged = True Then
        'change the context information, and reload the data if needed
        'save the context information to the instance sheet, and all other Args
        Call SaveFormData
        ' now disable the registration of data changing
        tl_tm_FormCtrl.EnableCtrl = False
        Call ReadSheet_SetForm 'context had changed, so reload values
        tl_tm_FormCtrl.EnableCtrl = True
    End If
    ' Validate the Data
    If ValData(TL_C_VALDATAMODENORMAL) Then
        'clear the datachanged flag
        tl_tm_FormCtrl.ChangedStatus = False
        ApplyButton.enabled = False
        tl_tm_FormCtrl.ButtonEnabled = ApplyButton.enabled
        If (tl_tm_BookIsValid = True) Then RunButton.enabled = True
    End If
    If tl_tm_FormCtrl.ContextChanged Then tl_tm_FormCtrl.ContextChanged = False
    'Save the Values of the form
    Call SaveFormData
End Sub
Private Sub CancelButton_Click()
    'Eliminate the form from the screen display
    Unload Me
End Sub
'Fix for tersw00072341 - Added this for the newly added edit box in Instance Editor Template.
Private Sub CommentsTextBox_Change()
 Call ControlResponse(CommentsTextBox)
End Sub
'Fix for TERSW00059979 - Added a text box to PowerSupply_T template
Private Sub DpsPrecondpatClampBox_Change()
    '------------------------------------------------------------------------------------
    'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
    'IG-XL Ver3.40.10ê¢ë„Ç≈ÇÕPowerSupply_TÇ≈tl_tm_ParDpsprecondpatClampÇ
    'ÉTÉ|Å[ÉgÇµÇƒÇ¢Ç»Ç¢ÇΩÇﬂÉRÉÅÉìÉgÉAÉEÉg
' Call ControlResponse(DpsPrecondpatClampBox)
    '------------------------------------------------------------------------------------
End Sub

Private Sub ExitButton_Click()
    'prepare to exit the userform
    If tl_tm_FormCtrl.ChangedStatus = True Then
        'Save the Values of the form
        Call SaveFormData
    End If
    Unload Me
End Sub

' Fix for TERW00074157
Private Sub HspGrpAndSetButton_Click()
Call ControlResponse(HspGrpAndSetButton, TL_C_PATSELFRMSTR, TL_C_PATSELAVLSTR, TL_C_PATSELSELSTR)
End Sub

' Fix for TERW00074157
Private Sub PcpGrpAndSetButton_Click()
Call ControlResponse(PcpGrpAndSetButton, TL_C_PATSELFRMSTR, TL_C_PATSELAVLSTR, TL_C_PATSELSELSTR)
End Sub

Private Sub RunButton_Click()
    If (tl_tm_BookIsValid = True) And (ApplyButton.enabled = False) Then
        Call tl_fs_RunIeInstance
    End If
End Sub

Private Sub SerializeMeasCheckBox_Click()
   ApplyButton.enabled = tl_fs_HandleCheck(SerializeMeasCheckBox, tl_tm_FormCtrl)
   tl_tm_ParSerializeMeas.ValueChanged = True
   tl_tm_ParSerializeMeas.ParameterValue = SerializeMeasCheckBox.Value
   If (SerializeMeasCheckBox.Value = True) Then
        tl_tm_ParSerializeMeas.ParameterStr = 1
   Else
        tl_tm_ParSerializeMeas.ParameterStr = 0
   End If
End Sub

Private Sub SerializeMeasFBox_Change()
    Call ControlResponse(SerializeMeasFBox)
End Sub

Private Sub SerializeMeasFInputBox_Change()
    Call ControlResponse(SerializeMeasFInputBox)
End Sub

Private Sub UserForm_Terminate()
    Call tl_tm_CleanUp
End Sub
Private Sub HelpButton_Click()
    'Run the help file by invoking the the help system with the files created by the Documentation group.
    Application.Help TL_C_HELPFILE, tl_tm_InstanceEditor.HelpValue
End Sub




'***************************************************
'NEXT 7 HANDLE INSTANCE-GROUP ORIENTED BUTTON CLICKS
'***************************************************
Private Sub GroupNumBox_Change()
    Dim MemCnt As Integer
    Dim MemNum  As Integer
    Dim RetBool As Boolean
    If tl_tm_OpMode = TL_C_INTERACTIVE Then
    
        Call tl_fs_MemberCount(MemCnt, MemNum)
        GroupCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(MemCnt)
        If (tl_tm_OpMode <> TL_C_BACKGROUND) And (Trim(GroupNumBox.Value) <> TL_C_EMPTYSTR) Then
            If IsNumeric(GroupNumBox.Value) And (CInt(val(GroupNumBox.Value)) = val(GroupNumBox.Value)) Then
                If ((CInt(GroupNumBox.Value) >= 0) And (CInt(GroupNumBox.Value) <= (MemCnt))) Then
                    'set the member number
                    RetBool = tl_dt_IESetActiveMember(CInt(GroupNumBox.Value))
                    
                    'get the data for that member
                    tl_tm_OpMode = TL_C_BACKGROUND
                    Call ReadSheet_SetForm
                    tl_tm_OpMode = TL_C_INTERACTIVE
                    
                    tl_tm_OpMode = TL_C_BACKGROUND
                    Call ValData(TL_C_VALDATAMODENOSTOP)
                    tl_tm_OpMode = TL_C_INTERACTIVE
                Else
                    tl_tm_OpMode = TL_C_BACKGROUND
                    GroupNumBox.Value = MemNum 'if a number out of range is entered, then reset number
                    tl_tm_OpMode = TL_C_INTERACTIVE
                End If
            Else
                tl_tm_OpMode = TL_C_BACKGROUND
                GroupNumBox.Value = MemNum 'if something is not a number, then reset number
                tl_tm_OpMode = TL_C_INTERACTIVE
            End If
        End If
        Call ControlButtons
    End If
End Sub
Private Sub GroupNumSpinButton_SpinDown()
    If CInt(GroupNumBox.Value) >= 1 Then
        GroupNumBox.Value = CInt(GroupNumBox.Value) - 1
    End If
End Sub
Private Sub GroupNumSpinButton_SpinUp()
    Dim MemCnt As Integer
    Dim MemNum As Integer
    Call tl_fs_MemberCount(MemCnt, MemNum)
    If CInt(GroupNumBox.Value) < MemCnt Then
        GroupNumBox.Value = CInt(GroupNumBox.Value) + 1
    End If
End Sub
Private Sub DeleteButton_Click()
    Dim MemCnt As Integer
    Dim MemNum  As Integer
    Call tl_fs_MemberCount(MemCnt, MemNum)
    If MemNum >= 1 Then

        Call tl_dt_IEDeleteMember
        Call tl_fs_MemberCount(MemCnt, MemNum)

        tl_tm_OpMode = TL_C_BACKGROUND
        Call ReadSheet_SetForm
        tl_tm_OpMode = TL_C_INTERACTIVE

        tl_tm_OpMode = TL_C_BACKGROUND
        Call ValData(TL_C_VALDATAMODENOSTOP)
        tl_tm_OpMode = TL_C_INTERACTIVE

        GroupCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(MemCnt)
        
        tl_tm_OpMode = TL_C_BACKGROUND
        GroupNumBox.Value = MemNum
        tl_tm_OpMode = TL_C_INTERACTIVE
    End If
    Call ControlButtons
End Sub
Private Sub AddButton_Click()
    Dim MemCnt As Integer
    Dim MemNum  As Integer
    ApplyButton.enabled = True
    Call tl_dt_IEInsertMember
    Call tl_fs_MemberCount(MemCnt, MemNum)
    
    tl_tm_OpMode = TL_C_BACKGROUND
    Call ReadSheet_SetForm
    tl_tm_OpMode = TL_C_INTERACTIVE

    tl_tm_OpMode = TL_C_BACKGROUND
    Call ValData(TL_C_VALDATAMODENOSTOP)
    tl_tm_OpMode = TL_C_INTERACTIVE

    GroupCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(MemCnt)
    tl_tm_OpMode = TL_C_BACKGROUND
    GroupNumBox.Value = MemNum
    tl_tm_OpMode = TL_C_INTERACTIVE
    Call ControlButtons
End Sub
Private Sub CopyButton_Click()
    For Each tl_tm_ParThisPar In AllPars
        tl_tm_ParThisPar.CopyGroup
    Next
    If tl_tm_MultiItemMode = False Then
        PasteButton.enabled = True
    End If
End Sub
Private Sub PasteButton_Click()
    For Each tl_tm_ParThisPar In AllPars
        tl_tm_ParThisPar.PasteGroup
    Next
    
    For Each tl_tm_ParThisPar In AllPars
        If Not (tl_tm_ParThisPar.DataBox Is Nothing) Then
            With tl_tm_ParThisPar
                If TypeOf .DataBox Is ComboBox Then
                    If .DataBox.MatchRequired = True Then
                        .ParameterValue = .ParameterStr
                        Call tl_fs_SetComboData(tl_tm_ParThisPar, TL_C_DELIMITERSTD)
                    Else
                        Call tl_fs_SetComboData(tl_tm_ParThisPar, TL_C_DELIMITERSTD)
                    End If
                Else 'it is a textbox
                    If Left(.ParameterStr, 1) <> "=" Then
                        .ParameterValue = .ParameterStr
                    Else
                        .ParameterValue = CStr(tl_dt_IEGetFormulaValue(.ParameterStr))
                    End If
                    Call tl_fs_SetBox(tl_tm_ParThisPar)
                End If
            
            'here the eval box should be updated if the context had changed.
            If tl_tm_FormCtrl.ContextChanged = True Then Call .ChangeMade(tl_tm_FormCtrl.EnableCtrl, tl_tm_FormCtrl.ChangedStatus, tl_tm_OpMode)
            End With
        End If
    Next
    
    'Pattern Conditioning - CondPatPage,
    If tl_tm_InstanceEditor.CondPatPage Then
        Call tl_fs_Set2WaitFlags(tl_tm_ParFlags, A1CheckBox, A0CheckBox, AXCheckBox, _
            B1CheckBox, B0CheckBox, BXCheckBox, _
            C1CheckBox, C0CheckBox, CXCheckBox, _
            D1CheckBox, D0CheckBox, DXCheckBox)
    End If
    If tl_tm_InstanceEditor.PatFuncPage Then
        Call tl_fs_Set2WaitFlags(tl_tm_ParPatFlags, PatA1CheckBox, PatA0CheckBox, PatAXCheckBox, _
            PatB1CheckBox, PatB0CheckBox, PatBXCheckBox, _
            PatC1CheckBox, PatC0CheckBox, PatCXCheckBox, _
            PatD1CheckBox, PatD0CheckBox, PatDXCheckBox)
    End If
    
    'Ppmu - HiLoLimValid,
    If tl_tm_InstanceEditor.PpmuPages Then
        Call tl_fs_SetCheckBoxes(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParPpmuLimits, _
            tl_tm_ParPpmuHiLimSpec, tl_tm_ParPpmuLoLimSpec)
    End If
    'Bpmu - HiLoLimValid,
    If tl_tm_InstanceEditor.BpmuPages Then
        Call tl_fs_SetCheckBoxes(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParBpmuLimits, _
            tl_tm_ParBpmuHiLimSpec, tl_tm_ParBpmuLoLimSpec)
    End If
    'Dps - HiLoLimValid,
    If tl_tm_InstanceEditor.DpsPages Then
        Call tl_fs_SetCheckBoxes(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, _
            tl_tm_ParDpsHiLimSpec, tl_tm_ParDpsLoLimSpec)
    End If
    
    ApplyButton.enabled = True
End Sub
Sub ControlButtons()
    Dim MemCnt As Integer
    Dim MemNum As Integer

    Call tl_fs_MemberCount(MemCnt, MemNum)
    If MemCnt < 1 Then
        GroupNumSpinButton.enabled = False
        AddButton.enabled = False
        DeleteButton.enabled = False
        GroupNumBox.enabled = False
    Else
        GroupNumSpinButton.enabled = True
        AddButton.enabled = True
        DeleteButton.enabled = True
        GroupNumBox.enabled = True
    End If
    
End Sub







'***************************************************
'NEXT 5 HANDLE PATTERN DRIVER SETTINGS
'***************************************************
Private Sub FuncSetPassFailCombo_Change()
    Call ControlResponse(FuncSetPassFailCombo)
End Sub
Private Sub FuncPatListEditButton_Click()
    Call ControlResponse(FuncPatListEditButton, TL_C_PATSELFRMSTR, TL_C_PATSELAVLSTR, TL_C_PATSELSELSTR)
End Sub
Private Sub FuncPatternTextBox_Change()
    Call ControlResponse(FuncPatternTextBox)
End Sub
Private Sub FuncFileFinderButton_Click()
    Call ControlResponse(FuncFileFinderButton)
End Sub
Private Sub FuncRelayModeCombo_Change()
    Call ControlResponse(FuncRelayModeCombo)
End Sub










'***************************************************
'NEXT 10 HANDLE COMMON PPMU & BPMU INSTRUMENT SETTINGS
'***************************************************
Private Sub PmuForceCond1Combo_Change()
    Call ControlResponse(PmuForceCond1Combo)
End Sub
Private Sub PmuForceCond1EquationButton_Click()
    Call ControlResponse(PmuForceCond1EquationButton)
End Sub
Private Sub PmuHiLimEquationButton_Click()
    Call ControlResponse(PmuHiLimEquationButton)
End Sub
Private Sub PmuHiLimitCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleCheck(PmuHiLimitCheckBox, tl_tm_FormCtrl)
    If tl_tm_InstanceEditor.BpmuPages Then
        Call tl_fs_CheckBoxesResponse(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParBpmuLimits, tl_tm_FormCtrl)
    Else
        Call tl_fs_CheckBoxesResponse(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParPpmuLimits, tl_tm_FormCtrl)
    End If
End Sub
Private Sub PmuHiLimSpecCombo_Change()
    Call ControlResponse(PmuHiLimSpecCombo)
End Sub
Private Sub PmuLoLimEquationButton_Click()
    Call ControlResponse(PmuLoLimEquationButton)
End Sub
Private Sub PmuLoLimitCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleCheck(PmuLoLimitCheckBox, tl_tm_FormCtrl)
    If tl_tm_InstanceEditor.BpmuPages Then
        Call tl_fs_CheckBoxesResponse(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParBpmuLimits, tl_tm_FormCtrl)
    Else
        Call tl_fs_CheckBoxesResponse(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParPpmuLimits, tl_tm_FormCtrl)
    End If
End Sub
Private Sub PmuLoLimSpecCombo_Change()
    Call ControlResponse(PmuLoLimSpecCombo)
End Sub
Private Sub PmuPinListBox_Change()
    Call ControlResponse(PmuPinListBox, TL_C_PINSELFRMSTR, TL_C_PINSELAVLSTR, TL_C_PINSELSELSTR)
End Sub
Private Sub PmuPinListEditButton_Click()
    Call ControlResponse(PmuPinListEditButton, TL_C_PINLISTSELFRMSTR, TL_C_PINLISTSELAVLSTR, TL_C_PINLISTSELSELSTR)
End Sub



'***************************************************
'NEXT 11 HANDLE OPTIONAL BPMU INSTRUMENT SETTINGS
'***************************************************
Private Sub BpmuForceCond2Combo_Change()
    Call ControlResponse(BpmuForceCond2Combo)
End Sub
Private Sub BpmuForceCond2EquationButton_Click()
    Call ControlResponse(BpmuForceCond2EquationButton)
End Sub
Private Sub BpmuIRangeCombo_Change()
    Call ControlResponse(BpmuIRangeCombo)
End Sub
Private Sub BpmuMeasureModeCombo_Change()
    Dim LastIrange As String
    Dim LastVrange As String
    Dim TargetPos As Long
    Dim CommaPos As Long
    Call ControlResponse(BpmuMeasureModeCombo)
    LastIrange = tl_tm_ParBpmuIRange.ParameterValue
    LastVrange = tl_tm_ParBpmuVrange.ParameterValue
    If tl_tm_ParBpmuMeasureMode.ParameterStr = tl_GetIndexOf(TL_C_MMISTR) Then
        tl_tm_ParBpmuIRange.ValueChoices = TheHdw.BPMU.MeasIRangeList
        tl_tm_ParBpmuVrange.ValueChoices = TheHdw.BPMU.ForceVRangeList
    Else
        tl_tm_ParBpmuIRange.ValueChoices = TheHdw.BPMU.ForceIRangeList
        tl_tm_ParBpmuVrange.ValueChoices = TheHdw.BPMU.MeasVRangeList
    End If
    Call tl_fs_RemoveComboPullDowns(tl_tm_ParBpmuIRange)
    Call tl_fs_SetComboPullDowns(tl_tm_ParBpmuIRange, TL_C_DELIMITERSTD)
    If LastIrange <> TL_C_EMPTYSTR Then
        TargetPos = InStr(tl_tm_ParBpmuIRange.ValueChoices, LastIrange & TL_C_DELIMITERRANGES)
        If TargetPos = 0 Then
            LastIrange = TL_C_SMARTRANGESTR
        Else
            CommaPos = InStr(TargetPos, tl_tm_ParBpmuIRange.ValueChoices, TL_C_DELIMITERSTD)
            LastIrange = Mid(tl_tm_ParBpmuIRange.ValueChoices, TargetPos + 2, CommaPos - TargetPos - 2)
        End If
        tl_tm_ParBpmuIRange.DataBox.Text = LastIrange
    End If
    Call tl_fs_RemoveComboPullDowns(tl_tm_ParBpmuVrange)
    Call tl_fs_SetComboPullDowns(tl_tm_ParBpmuVrange, TL_C_DELIMITERSTD)
    If LastVrange <> TL_C_EMPTYSTR Then
        TargetPos = InStr(tl_tm_ParBpmuVrange.ValueChoices, LastVrange & TL_C_DELIMITERRANGES)
        If TargetPos = 0 Then
            LastVrange = TL_C_SMARTRANGESTR
        Else
            CommaPos = InStr(TargetPos, tl_tm_ParBpmuVrange.ValueChoices, TL_C_DELIMITERSTD)
            LastVrange = Mid(tl_tm_ParBpmuVrange.ValueChoices, TargetPos + 2, CommaPos - TargetPos - 2)
        End If
        tl_tm_ParBpmuVrange.DataBox.Text = LastVrange
    End If
End Sub
Private Sub BpmuRelayModeCombo_Change()
    Call ControlResponse(BpmuRelayModeCombo)
End Sub
Private Sub BpmuSettlingTimeBox_Change()
    Call ControlResponse(BpmuSettlingTimeBox)
End Sub
Private Sub BpmuGangCombo_Change()
    Call ControlResponse(BpmuGangCombo)
End Sub
Private Sub BpmuSamplesBox_Change()
    Call ControlResponse(BpmuSamplesBox)
    If Trim(tl_tm_ParBpmuSamplingTime.DataBox.Text) <> TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamplingTime)
        Call tl_fs_MakeGray(tl_tm_ParBpmuSamples)
    ElseIf Trim(tl_tm_ParBpmuSamples.DataBox.Text) = TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamples)
    Else
        Call tl_fs_MakeGray(tl_tm_ParBpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamples)
    End If
End Sub
Private Sub BpmuSamplingTimeBox_Change()
    Call ControlResponse(BpmuSamplingTimeBox)
    If Trim(tl_tm_ParBpmuSamplingTime.DataBox.Text) <> TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamplingTime)
        Call tl_fs_MakeGray(tl_tm_ParBpmuSamples)
    ElseIf Trim(tl_tm_ParBpmuSamples.DataBox.Text) = TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamples)
    Else
        Call tl_fs_MakeGray(tl_tm_ParBpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParBpmuSamples)
    End If
End Sub
Private Sub BpmuClampBox_Change()
    Call ControlResponse(BpmuClampBox)
End Sub
Private Sub BpmuVRangeCombo_Change()
    Call ControlResponse(BpmuVRangeCombo)
End Sub









'***************************************************
'NEXT 5 HANDLE DPS TEST ITEM CONTROLS
'***************************************************
Private Sub DpsTestControlCombo_Change()
    Call ControlResponse(DpsTestControlCombo)
    If tl_tm_ParDpsTestControl.ParameterStr <> tl_GetIndexOf(TL_C_TCNORMSTR) Then
        'enable the multiple-item feature
        tl_tm_MultiItemMode = True
        ItemNumLabel.Visible = True
        ItemNumBox.Visible = True
        ItemNumSpinButton.Visible = True
        ItemCountBox.Visible = True
        ItemAddButton.Visible = True
        ItemDeleteButton.Visible = True
        
        'itemize the data items, if needed
        With tl_tm_ParDpsHiLimSpec
            If tl_tm_GetItemCnt(.ParameterItemsStr) = 0 Then
                .ParameterItemsStr = TL_C_DELIMITERGROUPS & .ParameterStr & TL_C_DELIMITERVALUE & .ParameterValue & TL_C_DELIMITERGROUPS
            End If
        End With
        With tl_tm_ParDpsLoLimSpec
            If tl_tm_GetItemCnt(.ParameterItemsStr) = 0 Then
                .ParameterItemsStr = TL_C_DELIMITERGROUPS & .ParameterStr & TL_C_DELIMITERVALUE & .ParameterValue & TL_C_DELIMITERGROUPS
            End If
        End With
        With tl_tm_ParDpsSettlingTime
            If tl_tm_GetItemCnt(.ParameterItemsStr) = 0 Then
                .ParameterItemsStr = TL_C_DELIMITERGROUPS & .ParameterStr & TL_C_DELIMITERVALUE & .ParameterValue & TL_C_DELIMITERGROUPS
            End If
        End With
        With tl_tm_ParDpsLimits
            If tl_tm_GetItemCnt(.ParameterItemsStr) = 0 Then
                .ParameterItemsStr = TL_C_DELIMITERGROUPS & .ParameterStr & TL_C_DELIMITERVALUE & .ParameterValue & TL_C_DELIMITERGROUPS
            End If
        End With
        
        tl_tm_MultiItemCnt = tl_tm_GetItemCnt(tl_tm_ParDpsHiLimSpec.ParameterItemsStr)
        tl_tm_MultiItemNum = 0

        ItemCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(tl_tm_MultiItemCnt)
        tl_tm_OpMode = TL_C_BACKGROUND
        ItemNumBox.Text = CStr(tl_tm_MultiItemNum)
        tl_tm_OpMode = TL_C_INTERACTIVE
        
    Else
        'disable the multiple-item feature - account for the Item capability
        tl_tm_MultiItemMode = False
        ItemNumLabel.Visible = False
        ItemNumBox.Visible = False
        ItemNumSpinButton.Visible = False
        ItemCountBox.Visible = False
        ItemAddButton.Visible = False
        ItemDeleteButton.Visible = False
        
    End If
    tl_tm_ParDpsHiLimSpec.ValueChanged = True
    tl_tm_ParDpsLoLimSpec.ValueChanged = True
    tl_tm_ParDpsSettlingTime.ValueChanged = True
    tl_tm_ParDpsLimits.ValueChanged = True
    
End Sub
Private Sub ItemAddButton_Click()
    Dim temp As String
    Dim temp2 As String
    ApplyButton.enabled = True
    tl_tm_MultiItemNum = tl_tm_MultiItemNum + 1
    
    With tl_tm_ParDpsHiLimSpec
        temp = .ParameterItemsStr
        temp2 = .defaultvalue & TL_C_DELIMITERVALUE
        Call tl_tm_AddItemStr(temp, temp2, tl_tm_MultiItemNum)
        .ParameterItemsStr = temp
        .DataBox.Text = .defaultvalue
    End With
    With tl_tm_ParDpsLoLimSpec
        temp = .ParameterItemsStr
        temp2 = .defaultvalue & TL_C_DELIMITERVALUE
        Call tl_tm_AddItemStr(temp, temp2, tl_tm_MultiItemNum)
        .ParameterItemsStr = temp
        .DataBox.Text = .defaultvalue
    End With
    With tl_tm_ParDpsSettlingTime
        temp = .ParameterItemsStr
        temp2 = .defaultvalue & TL_C_DELIMITERVALUE
        Call tl_tm_AddItemStr(temp, temp2, tl_tm_MultiItemNum)
        .ParameterItemsStr = temp
        .DataBox.Text = .defaultvalue
    End With
    With tl_tm_ParDpsLimits
        temp = .ParameterItemsStr
        temp2 = .defaultvalue & TL_C_DELIMITERVALUE & .defaultvalue
        Call tl_tm_AddItemStr(temp, temp2, tl_tm_MultiItemNum)
        .ParameterItemsStr = temp
        .ParameterStr = .defaultvalue
        .ValueChanged = True
    End With
    Call tl_fs_SetCheckBoxes(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, _
        tl_tm_ParDpsHiLimSpec, tl_tm_ParDpsLoLimSpec)

    tl_tm_OpMode = TL_C_BACKGROUND
    Call ValData(TL_C_VALDATAMODENOSTOP)
    ItemNumBox.Value = tl_tm_MultiItemNum
    tl_tm_OpMode = TL_C_INTERACTIVE
    
    tl_tm_MultiItemCnt = tl_tm_GetItemCnt(tl_tm_ParDpsHiLimSpec.ParameterItemsStr)
    ItemCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(tl_tm_MultiItemCnt)
End Sub
Private Sub ItemDeleteButton_Click()
    Dim lngX As Long
    Dim DecItemNum As Boolean
    Dim temp As String
    tl_tm_MultiItemCnt = tl_tm_GetItemCnt(tl_tm_ParDpsHiLimSpec.ParameterItemsStr)
    DecItemNum = False
    If tl_tm_MultiItemCnt = CLng(ItemNumBox.Value) Then
        lngX = tl_tm_MultiItemCnt
        If tl_tm_MultiItemCnt > 0 Then DecItemNum = True
        tl_tm_MultiItemNum = tl_tm_MultiItemNum - 1
    Else
        lngX = tl_tm_MultiItemCnt - 1
        If lngX < 0 Then lngX = 0
    End If
    
    temp = tl_tm_ParDpsHiLimSpec.ParameterItemsStr
    Call tl_tm_RemoveItemStr(temp, CLng(ItemNumBox.Value))
    tl_tm_ParDpsHiLimSpec.ParameterItemsStr = temp
    temp = tl_tm_ParDpsLoLimSpec.ParameterItemsStr
    Call tl_tm_RemoveItemStr(temp, CLng(ItemNumBox.Value))
    tl_tm_ParDpsLoLimSpec.ParameterItemsStr = temp
    temp = tl_tm_ParDpsSettlingTime.ParameterItemsStr
    Call tl_tm_RemoveItemStr(temp, CLng(ItemNumBox.Value))
    tl_tm_ParDpsSettlingTime.ParameterItemsStr = temp
    temp = tl_tm_ParDpsLimits.ParameterItemsStr
    Call tl_tm_RemoveItemStr(temp, CLng(ItemNumBox.Value))
    tl_tm_ParDpsLimits.ParameterItemsStr = temp
    
    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsHiLimSpec.ParameterItemsStr, tl_tm_MultiItemNum)
    tl_tm_ParDpsHiLimSpec.ParameterStr = tl_tm_GetItemForm(temp)
    tl_tm_ParDpsHiLimSpec.DataBox.Text = tl_tm_ParDpsHiLimSpec.ParameterStr

    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsLoLimSpec.ParameterItemsStr, tl_tm_MultiItemNum)
    tl_tm_ParDpsLoLimSpec.ParameterStr = tl_tm_GetItemForm(temp)
    tl_tm_ParDpsLoLimSpec.DataBox.Text = tl_tm_ParDpsLoLimSpec.ParameterStr
    
    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsSettlingTime.ParameterItemsStr, tl_tm_MultiItemNum)
    tl_tm_ParDpsSettlingTime.ParameterStr = tl_tm_GetItemForm(temp)
    tl_tm_ParDpsSettlingTime.DataBox.Text = tl_tm_ParDpsSettlingTime.ParameterStr
    
    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsLimits.ParameterItemsStr, tl_tm_MultiItemNum)
    tl_tm_ParDpsLimits.ParameterStr = tl_tm_GetItemForm(temp)
    Call tl_fs_SetCheckBoxes(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, _
        tl_tm_ParDpsHiLimSpec, tl_tm_ParDpsLoLimSpec)

    tl_tm_OpMode = TL_C_BACKGROUND
    Call ValData(TL_C_VALDATAMODENOSTOP)
    tl_tm_OpMode = TL_C_INTERACTIVE

    tl_tm_MultiItemCnt = tl_tm_GetItemCnt(tl_tm_ParDpsHiLimSpec.ParameterItemsStr)
    ItemCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(tl_tm_MultiItemCnt)
    
    If DecItemNum Then
        tl_tm_OpMode = TL_C_BACKGROUND
        ItemNumBox.Value = tl_tm_MultiItemNum
        tl_tm_OpMode = TL_C_INTERACTIVE
    End If
End Sub
Private Sub ItemNumBox_Change()
    Dim temp As String
    If tl_tm_OpMode = TL_C_INTERACTIVE Then
    
        tl_tm_MultiItemCnt = tl_tm_GetItemCnt(tl_tm_ParDpsHiLimSpec.ParameterItemsStr)
        If (Trim(ItemNumBox.Value) <> TL_C_EMPTYSTR) Then
            If IsNumeric(ItemNumBox.Value) And (CInt(val(ItemNumBox.Value)) = val(ItemNumBox.Value)) Then
                If ((CInt(ItemNumBox.Value) >= 0) And (CInt(ItemNumBox.Value) <= (tl_tm_MultiItemCnt))) Then
                    tl_tm_MultiItemNum = ItemNumBox.Value
                    
                    'get the data for that member
                    tl_tm_OpMode = TL_C_BACKGROUND
                    
                    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsHiLimSpec.ParameterItemsStr, tl_tm_MultiItemNum)
                    tl_tm_ParDpsHiLimSpec.ParameterStr = tl_tm_GetItemForm(temp)
                    tl_tm_ParDpsHiLimSpec.DataBox.Text = tl_tm_ParDpsHiLimSpec.ParameterStr
                    
                    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsLoLimSpec.ParameterItemsStr, tl_tm_MultiItemNum)
                    tl_tm_ParDpsLoLimSpec.ParameterStr = tl_tm_GetItemForm(temp)
                    tl_tm_ParDpsLoLimSpec.DataBox.Text = tl_tm_ParDpsLoLimSpec.ParameterStr
                    
                    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsSettlingTime.ParameterItemsStr, tl_tm_MultiItemNum)
                    tl_tm_ParDpsSettlingTime.ParameterStr = tl_tm_GetItemForm(temp)
                    tl_tm_ParDpsSettlingTime.DataBox.Text = tl_tm_ParDpsSettlingTime.ParameterStr

                    temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsLimits.ParameterItemsStr, tl_tm_MultiItemNum)
                    tl_tm_ParDpsLimits.ParameterStr = tl_tm_GetItemForm(temp)
                    Call tl_fs_SetCheckBoxes(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, _
                        tl_tm_ParDpsHiLimSpec, tl_tm_ParDpsLoLimSpec)
                    
                    tl_tm_OpMode = TL_C_INTERACTIVE
                    
                    tl_tm_OpMode = TL_C_BACKGROUND
                    Call ValData(TL_C_VALDATAMODENOSTOP)
                    tl_tm_OpMode = TL_C_INTERACTIVE
                Else
                    tl_tm_OpMode = TL_C_BACKGROUND
                    ItemNumBox.Value = tl_tm_MultiItemNum 'if a number out of range is entered, then reset number
                    tl_tm_OpMode = TL_C_INTERACTIVE
                End If
            Else
                tl_tm_OpMode = TL_C_BACKGROUND
                ItemNumBox.Value = tl_tm_MultiItemNum 'if something is not a number, then reset number
                tl_tm_OpMode = TL_C_INTERACTIVE
            End If
        End If
    End If
End Sub
Private Sub ItemNumSpinButton_SpinDown()
    If CInt(ItemNumBox.Value) >= 1 Then
        tl_tm_MultiItemNum = tl_tm_MultiItemNum - 1
        ItemNumBox.Value = tl_tm_MultiItemNum
    End If
End Sub
Private Sub ItemNumSpinButton_SpinUp()
    If CInt(ItemNumBox.Value) < tl_tm_GetItemCnt(tl_tm_ParDpsHiLimSpec.ParameterItemsStr) Then
        tl_tm_MultiItemNum = tl_tm_MultiItemNum + 1
        ItemNumBox.Value = tl_tm_MultiItemNum
    End If
End Sub







'***************************************************
'NEXT 19 HANDLE DPS INSTRUMENT SETTINGS
'***************************************************
Private Sub DpsPowerPinListEditButton_Click()
    Call ControlResponse(DpsPowerPinListEditButton, TL_C_PINLISTSELFRMSTR, TL_C_PINLISTSELAVLSTR, TL_C_PINLISTSELSELSTR)
End Sub
Private Sub DpsMainForceCondCombo_Change()
    Call ControlResponse(DpsMainForceCondCombo)
End Sub
Private Sub DpsMainForceCondEquationButton_Click()
    Call ControlResponse(DpsMainForceCondEquationButton)
End Sub
Private Sub DpsAltForceCondCombo_Change()
    Call ControlResponse(DpsAltForceCondCombo)
End Sub
Private Sub DpsAltForceCondEquationButton_Click()
    Call ControlResponse(DpsAltForceCondEquationButton)
End Sub
Private Sub DpsHiLimEquationButton_Click()
    Call ControlResponse(DpsHiLimEquationButton)
End Sub
Private Sub DpsHiLimitCheckBox_Click()
    Dim temp As String
    ApplyButton.enabled = tl_fs_HandleCheck(DpsHiLimitCheckBox, tl_tm_FormCtrl)
    Call tl_fs_CheckBoxesResponse(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, tl_tm_FormCtrl)
    If InStr(tl_tm_ParDpsLimits.ParameterItemsStr, TL_C_DELIMITERGROUPS) <> 0 Then
        temp = tl_tm_ParDpsLimits.ParameterItemsStr
        Call tl_tm_InsertItemStr(temp, tl_tm_ParDpsLimits.ParameterStr & TL_C_DELIMITERVALUE & tl_tm_ParDpsLimits.ParameterStr, tl_tm_MultiItemNum, tl_tm_MultiItemCnt)
        tl_tm_ParDpsLimits.ParameterItemsStr = temp
    End If
End Sub
Private Sub DpsHiLimSpecCombo_Change()
    Call ControlResponse(DpsHiLimSpecCombo)
End Sub
Private Sub DpsLoLimEquationButton_Click()
    Call ControlResponse(DpsLoLimEquationButton)
End Sub
Private Sub DpsLoLimitCheckBox_Click()
    Dim temp As String
    ApplyButton.enabled = tl_fs_HandleCheck(DpsLoLimitCheckBox, tl_tm_FormCtrl)
    Call tl_fs_CheckBoxesResponse(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, tl_tm_FormCtrl)
    If InStr(tl_tm_ParDpsLimits.ParameterItemsStr, TL_C_DELIMITERGROUPS) <> 0 Then
        temp = tl_tm_ParDpsLimits.ParameterItemsStr
        Call tl_tm_InsertItemStr(temp, tl_tm_ParDpsLimits.ParameterStr & TL_C_DELIMITERVALUE & tl_tm_ParDpsLimits.ParameterStr, tl_tm_MultiItemNum, tl_tm_MultiItemCnt)
        tl_tm_ParDpsLimits.ParameterItemsStr = temp
    End If
End Sub
Private Sub DpsLoLimSpecCombo_Change()
    Call ControlResponse(DpsLoLimSpecCombo)
End Sub
Private Sub DpsIRangeCombo_Change()
    Call ControlResponse(DpsIRangeCombo)
End Sub
Private Sub DpsSettlingTimeBox_Change()
    Call ControlResponse(DpsSettlingTimeBox)
End Sub
Private Sub DpsPowerPinListBox_Change()
    Dim temp As String
    Dim QueGroup As String
    Dim CountA As Long
    Dim CountB As Long
    Dim lngX As Long
    Dim PwrArr() As String
    Call ControlResponse(DpsPowerPinListBox, TL_C_PINSELFRMSTR, TL_C_PINSELAVLSTR, TL_C_PINSELSELSTR)
    temp = tl_dt_IEGetPinGroupList(TL_DT_PIN_PWR)
    'the current limit must be set to be that which is the current of the single or least-ganged power supply pin.
    CountA = 99
    Call tl_tm_StrToArr(tl_tm_ParDpsPowerPin.ParameterValue, PwrArr)
    lngX = 0
    If UBound(PwrArr) = 0 Then CountA = 1
    Do While (lngX < UBound(PwrArr)) And (CountA <> 1)
        lngX = lngX + 1
        If InStr(1, PwrArr(lngX), temp, vbTextCompare) = 0 Then
            'pin is not a ganged power supply
            CountA = 1
        Else
            'pin is a ganged power supply
            ' see how many pins are actually ganged together.
            CountB = tl_dt_IEDecomposePinList(PwrArr(lngX), QueGroup)
            If CountB < CountA Then CountA = CountB
        End If
    Loop
    tl_tm_ParDpsIRange.RangeLimits = TL_C_LRANGEINDEX & TL_C_DELIMITERRANGES & "-0.1" & TL_C_DELIMITERSTD & TL_C_HRANGEINDEX & TL_C_DELIMITERRANGES & CStr(CountA * tl_tm_GetRangeVal(tl_tm_ParDpsIRange.ParameterStr, tl_tm_ParDpsIRange.ValueChoices))
End Sub
Private Sub DpsMainOrAltCombo_Change()
    Call ControlResponse(DpsMainOrAltCombo)
    If InStr(TL_C_MOAASTR, tl_tm_ParDpsMAINORALT.ParameterValue & TL_C_DELIMITERRANGES) Then
        Call tl_fs_MakeBlack(tl_tm_ParDpsAltForceCond)
        Call tl_fs_MakeGray(tl_tm_ParDpsMainForceCond)
    End If
    If InStr(TL_C_MOAMSTR, tl_tm_ParDpsMAINORALT.ParameterValue & TL_C_DELIMITERRANGES) Then
        Call tl_fs_MakeGray(tl_tm_ParDpsAltForceCond)
        Call tl_fs_MakeBlack(tl_tm_ParDpsMainForceCond)
    End If
    If InStr(TL_C_MOALSTR, tl_tm_ParDpsMAINORALT.ParameterValue & TL_C_DELIMITERRANGES) Then
        Call tl_fs_MakeGray(tl_tm_ParDpsAltForceCond)
        Call tl_fs_MakeGray(tl_tm_ParDpsMainForceCond)
    End If
End Sub
Private Sub DpsSamplesBox_Change()
    Call ControlResponse(DpsSamplesBox)
    If Trim(tl_tm_ParDpsSamples.DataBox.Text) <> TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamples)
'        Call tl_fs_MakeGray(tl_tm_ParDpsSamplingTime)
    ElseIf Trim(tl_tm_ParDpsSamplingTime.DataBox.Text) = TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamples)
    Else
        Call tl_fs_MakeGray(tl_tm_ParDpsSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamples)
    End If
End Sub
Private Sub DpsSamplingTimeBox_Change()
    Call ControlResponse(DpsSamplingTimeBox)
    If Trim(tl_tm_ParDpsSamplingTime.DataBox.Text) <> TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamplingTime)
        Call tl_fs_MakeGray(tl_tm_ParDpsSamples)
    ElseIf Trim(tl_tm_ParDpsSamples.DataBox.Text) = TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamples)
    Else
        Call tl_fs_MakeGray(tl_tm_ParDpsSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParDpsSamples)
    End If
End Sub
Private Sub DpsClampBox_Change()
    Call ControlResponse(DpsClampBox)
End Sub
Private Sub DpsRelayModeCombo_Change()
    Call ControlResponse(DpsRelayModeCombo)
End Sub







'***************************************************
'NEXT 11 HANDLE OPTIONAL PPMU INSTRUMENT SETTINGS
'***************************************************
Private Sub PpmuFloadCombo_Change()
    Call ControlResponse(PpmuFloadCombo)
End Sub
Private Sub PpmuForceCond2Combo_Change()
    Call ControlResponse(PpmuForceCond2Combo)
End Sub
Private Sub PpmuForceCond2EquationButton_Click()
    Call ControlResponse(PpmuForceCond2EquationButton)
End Sub
Private Sub PpmuIRangeCombo_Change()
    Call ControlResponse(PpmuIRangeCombo)
End Sub
Private Sub PpmuMeasureModeCombo_Change()
    Dim LastIrange As String
    Dim TargetPos As Long
    Dim CommaPos As Long
    Call ControlResponse(PpmuMeasureModeCombo)
    LastIrange = tl_tm_ParPpmuIRange.ParameterValue
    If tl_tm_ParPpmuMeasureMode.ParameterStr = tl_GetIndexOf(TL_C_MMISTR) Then
        tl_tm_ParPpmuIRange.ValueChoices = TheHdw.PPMU.MeasIRangeList
    Else
        tl_tm_ParPpmuIRange.ValueChoices = TheHdw.PPMU.ForceIRangeList
    End If
    Call tl_fs_RemoveComboPullDowns(tl_tm_ParPpmuIRange)
    Call tl_fs_SetComboPullDowns(tl_tm_ParPpmuIRange, TL_C_DELIMITERSTD)
    If LastIrange <> TL_C_EMPTYSTR Then
        TargetPos = InStr(tl_tm_ParPpmuIRange.ValueChoices, LastIrange & TL_C_DELIMITERRANGES)
        If TargetPos = 0 Then
            LastIrange = TL_C_SMARTRANGESTR
        Else
            CommaPos = InStr(TargetPos, tl_tm_ParPpmuIRange.ValueChoices, TL_C_DELIMITERSTD)
            LastIrange = Mid(tl_tm_ParPpmuIRange.ValueChoices, TargetPos + 2, CommaPos - TargetPos - 2)
        End If
        tl_tm_ParPpmuIRange.DataBox.Text = LastIrange
    End If
End Sub
Private Sub PpmuRelayModeCombo_Change()
    Call ControlResponse(PpmuRelayModeCombo)
End Sub
Private Sub PpmuSettlingTimeBox_Change()
    Call ControlResponse(PpmuSettlingTimeBox)
End Sub
Private Sub PpmuSamplingTimeBox_Change()
    Call ControlResponse(PpmuSamplingTimeBox)
    If Trim(tl_tm_ParPpmuSamplingTime.DataBox.Text) <> TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamplingTime)
        Call tl_fs_MakeGray(tl_tm_ParPpmuSamples)
    ElseIf Trim(tl_tm_ParPpmuSamples.DataBox.Text) = TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamples)
    Else
        Call tl_fs_MakeGray(tl_tm_ParPpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamples)
    End If
End Sub
Private Sub PpmuSamplesBox_Change()
    Call ControlResponse(PpmuSamplesBox)
    If Trim(tl_tm_ParPpmuSamplingTime.DataBox.Text) <> TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamplingTime)
        Call tl_fs_MakeGray(tl_tm_ParPpmuSamples)
    ElseIf Trim(tl_tm_ParPpmuSamples.DataBox.Text) = TL_C_EMPTYSTR Then
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamples)
    Else
        Call tl_fs_MakeGray(tl_tm_ParPpmuSamplingTime)
        Call tl_fs_MakeBlack(tl_tm_ParPpmuSamples)
    End If
End Sub
Private Sub PpmuHClampBox_Change()
    Call ControlResponse(PpmuHClampBox)
End Sub
Private Sub PpmuLClampBox_Change()
    Call ControlResponse(PpmuLClampBox)
End Sub





'***************************************************
'NEXT 16 HANDLE PATTERNS
'***************************************************
Private Sub PcpStopLabelBox_Change()
    Call ControlResponse(PcpStopLabelBox)
End Sub
Private Sub PcpStartLabelBox_Change()
    Call ControlResponse(PcpStartLabelBox)
End Sub
Private Sub FlagWaitTimeBox_Change()
    Call ControlResponse(FlagWaitTimeBox)
End Sub
Private Sub HoldStatePatCombo_Change()
    ' label box is grayed if the pattern is pattern set
    Dim PatSet As String
    PatSet = tl_dt_IEGetPatternSetList
    Call ControlResponse(HoldStatePatCombo)
    If InStr(PatSet, tl_tm_ParThisPar.DataBox) Then
       tl_tm_ParHspStartLabel.DataBox = ""
       tl_tm_ParHspStopLabel.DataBox = ""
       Call tl_fs_MakeGray(tl_tm_ParHspStartLabel)
       Call tl_fs_MakeGray(tl_tm_ParHspStopLabel)
    Else
       Call tl_fs_MakeBlack(tl_tm_ParHspStartLabel)
       Call tl_fs_MakeBlack(tl_tm_ParHspStopLabel)
    End If
End Sub
Private Sub PreCondPatCombo_Change()
    ' label box is grayed if the pattern is pattern set
    Dim PatSet As String
    PatSet = tl_dt_IEGetPatternSetList
    Call ControlResponse(PreCondPatCombo)
    If InStr(PatSet, tl_tm_ParThisPar.DataBox) Then
       tl_tm_ParPcpStartLabel.DataBox = ""
       tl_tm_ParPcpStopLabel.DataBox = ""
       Call tl_fs_MakeGray(tl_tm_ParPcpStartLabel)
       Call tl_fs_MakeGray(tl_tm_ParPcpStopLabel)
    Else
       Call tl_fs_MakeBlack(tl_tm_ParPcpStartLabel)
       Call tl_fs_MakeBlack(tl_tm_ParPcpStopLabel)
    End If
End Sub
Private Sub HspFinderButton_Click()
    Call ControlResponse(HspFinderButton)
End Sub
Private Sub PcpFinderButton_Click()
    Call ControlResponse(PcpFinderButton)
End Sub
Private Sub HspStartLabelBox_Change()
    Call ControlResponse(HspStartLabelBox)
End Sub
Private Sub HspStopLabelBox_Change()
    Call ControlResponse(HspStopLabelBox)
End Sub
Private Sub HspCheckPgCombo_Change()
    Call ControlResponse(HspCheckPgCombo)
End Sub
Private Sub PcpCheckPgCombo_Change()
    Call ControlResponse(PcpCheckPgCombo)
End Sub
Private Sub HspResumeCombo_Change()
    Call ControlResponse(HspResumeCombo)
End Sub
Private Sub A0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 1, A1CheckBox, A0CheckBox, AXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub A1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 1, A1CheckBox, A0CheckBox, AXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub AXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 1, A1CheckBox, A0CheckBox, AXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub B0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 2, B1CheckBox, B0CheckBox, BXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub B1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 2, B1CheckBox, B0CheckBox, BXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub BXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 2, B1CheckBox, B0CheckBox, BXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub C0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 3, C1CheckBox, C0CheckBox, CXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub C1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 3, C1CheckBox, C0CheckBox, CXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub CXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 3, C1CheckBox, C0CheckBox, CXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub D0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 4, D1CheckBox, D0CheckBox, DXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub D1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 4, D1CheckBox, D0CheckBox, DXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub DXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 4, D1CheckBox, D0CheckBox, DXCheckBox, tl_tm_ParFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub




'***************************************************
'NEXT 15 HANDLE PATTERN FUNCTION CONTROLS
'***************************************************
Private Sub PatA0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 1, PatA1CheckBox, PatA0CheckBox, PatAXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatA1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 1, PatA1CheckBox, PatA0CheckBox, PatAXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatAXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 1, PatA1CheckBox, PatA0CheckBox, PatAXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatB0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 2, PatB1CheckBox, PatB0CheckBox, PatBXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatB1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 2, PatB1CheckBox, PatB0CheckBox, PatBXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatBXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 2, PatB1CheckBox, PatB0CheckBox, PatBXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatC0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 3, PatC1CheckBox, PatC0CheckBox, PatCXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatC1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 3, PatC1CheckBox, PatC0CheckBox, PatCXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatCXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 3, PatC1CheckBox, PatC0CheckBox, PatCXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatD0CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(0, 4, PatD1CheckBox, PatD0CheckBox, PatDXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatD1CheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(1, 4, PatD1CheckBox, PatD0CheckBox, PatDXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatDXCheckBox_Click()
    ApplyButton.enabled = tl_fs_HandleFlagCheck(2, 4, PatD1CheckBox, PatD0CheckBox, PatDXCheckBox, tl_tm_ParPatFlags, tl_tm_FormCtrl, tl_tm_OpMode)
End Sub
Private Sub PatFunctionBox_Change()
    Call ControlResponse(PatFunctionBox)
End Sub
Private Sub PatFunctionInput_Change()
    Call ControlResponse(PatFunctionInput)
End Sub
Private Sub PatFlagWaitTimeBox_Change()
    Call ControlResponse(PatFlagWaitTimeBox)
End Sub
Private Sub ThreadingCombo_Change()
    Call ControlResponse(ThreadingCombo)
End Sub
Private Sub MatchAllSitesCombo_Change()
    Call ControlResponse(MatchAllSitesCombo)
End Sub








'***************************************************
'NEXT 9 HANDLE TIMING, LEVELS, EDGES SHEETS
'***************************************************
Private Sub AcCatCombo_Change()
    Call ControlResponse(AcCatCombo)
    Call ContextChanged
    If (JobData.AvailAcCat = TL_C_EMPTYSTR) And (tl_tm_ParAcCat.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParAcCat)
    End If
End Sub
Private Sub DcCatCombo_Change()
    Call ControlResponse(DcCatCombo)
    Call ContextChanged
    If (JobData.AvailDcCat = TL_C_EMPTYSTR) And (tl_tm_ParDcCat.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParDcCat)
    End If
End Sub
Private Sub AcSelCombo_Change()
    Call ControlResponse(AcSelCombo)
    Call ContextChanged
    If (JobData.AvailAcSel = TL_C_EMPTYSTR) And (tl_tm_ParAcSel.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParAcSel)
    End If
End Sub
Private Sub DcSelCombo_Change()
    Call ControlResponse(DcSelCombo)
    Call ContextChanged
    If (JobData.AvailDcSel = TL_C_EMPTYSTR) And (tl_tm_ParDcSel.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParDcSel)
    End If
End Sub
Private Sub EdgeSetCombo_Change()
    Call ControlResponse(EdgeSetCombo)
    Call ContextChanged
    ' dpl 03/22/02 CQ7536 - Leave the EdgeSet in Black when selecting an empty string.
    'If tl_tm_ParEdgeSet.ParameterStr = TL_C_EMPTYSTR Then
        ' If the selected Edge Set is an empty string, set the TimeSet elements 'Black', and the EdgeSet elements 'Gray'
    '    Call tl_fs_BlackGrayByMouse(tl_tm_ParTimeSet, tl_tm_ParEdgeSet)
    'Else
    '    Call tl_fs_MakeBlack(tl_tm_ParEdgeSet)
    'End If
    If (JobData.AvailEdgeSet = TL_C_EMPTYSTR) And (tl_tm_ParEdgeSet.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParEdgeSet)
    End If
End Sub
Private Sub EdgeSetCombo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If tl_tm_ParEdgeSet.DataBox.enabled = False Then Exit Sub
    If tl_tm_ParTimeSet.DataBox.enabled = False Then Exit Sub
    
    ' dpl 03/22/02 CQ7536 - Leave the EdgeSet in Black when selecting an empty string.
    'If tl_tm_ParEdgeSet.ParameterStr = TL_C_EMPTYSTR Then
        ' If the selected Edge Set is an empty string, set the TimeSet elements 'Black', and the EdgeSet elements 'Gray'
    '    Call tl_fs_BlackGrayByMouse(tl_tm_ParTimeSet, tl_tm_ParEdgeSet)
    'Else
    '    Call tl_fs_MakeBlack(tl_tm_ParEdgeSet)
    'End If
    If (JobData.AvailEdgeSet = TL_C_EMPTYSTR) And (tl_tm_ParEdgeSet.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParEdgeSet)
    End If
End Sub
Private Sub TimeSetCombo_Change()
    Call ControlResponse(TimeSetCombo)
    Call ContextChanged
    ' see if the timeset sheet chosen is timesetbasic
    If InStr(JobData.AvailTimeSetExtended, tl_tm_ParTimeSet.ParameterStr) = 0 Then
        'this is a timeset basic case
        Call tl_fs_MakeGray(tl_tm_ParEdgeSet)
    Else
        'this is a regular timeset case
        Call tl_fs_MakeBlack(tl_tm_ParEdgeSet)
    End If
    If (JobData.AvailTimeSetAll = TL_C_EMPTYSTR) And (tl_tm_ParTimeSet.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParTimeSet)
    End If
End Sub
Private Sub LevelsCombo_Change()
    Call ControlResponse(LevelsCombo)
    Call ContextChanged
    If (JobData.AvailLevels = TL_C_EMPTYSTR) And (tl_tm_ParLevels.ParameterStr = TL_C_EMPTYSTR) Then
        Call tl_fs_MakeGray(tl_tm_ParLevels)
    End If
End Sub
Private Sub ContextChanged()
    Dim AllSpecs As String
    Dim AllSpecEntries As String
    If tl_tm_FormCtrl.EnableCtrl = True Then
        tl_tm_FormCtrl.ContextChanged = True
    End If
    
    Call tl_fs_TemplateSpecsStrings(AllSpecs, AllSpecEntries)
    If (AllSpecs <> JobData.AllSpecs) Or (AllSpecEntries <> JobData.AllSpecEntries) Then
        JobData.AllSpecs = AllSpecs
        JobData.AllSpecEntries = AllSpecEntries
        For Each tl_tm_ParThisPar In AllPars
            If tl_tm_ParThisPar.VCisSpecEntries Then
                Call tl_fs_RemoveComboPullDowns(tl_tm_ParThisPar)
                tl_tm_ParThisPar.ValueChoices = JobData.AllSpecEntries
                Call tl_fs_SetComboPullDowns(tl_tm_ParThisPar, TL_C_DELIMITERSTD)
            End If
        Next tl_tm_ParThisPar
    End If
End Sub
Private Sub OverlayCombo_Change()
    Call ControlResponse(OverlayCombo)
End Sub






'***************************************************
'NEXT 6 HANDLE INTERPOSE FUNCTION NAMES
'***************************************************
Private Sub PostPatFBox_Change()
    Call ControlResponse(PostPatFBox)
End Sub
Private Sub PostTestBox_Change()
    Call ControlResponse(PostTestBox)
End Sub
Private Sub PrePatFBox_Change()
    Call ControlResponse(PrePatFBox)
End Sub
Private Sub PreTestBox_Change()
    Call ControlResponse(PreTestBox)
End Sub
Private Sub StartOfBodyBox_Change()
    Call ControlResponse(StartOfBodyBox)
End Sub
Private Sub EndOfBodyBox_Change()
    Call ControlResponse(EndOfBodyBox)
End Sub
'***************************************************
'NEXT 6 HANDLE INTERPOSE FUNCTION INPUT ARGUMENTS
'***************************************************
Private Sub StartOfBodyInputBox_Change()
    Call ControlResponse(StartOfBodyInputBox)
End Sub
Private Sub PrePatInputBox_Change()
    Call ControlResponse(PrePatInputBox)
End Sub
Private Sub PreTestInputBox_Change()
    Call ControlResponse(PreTestInputBox)
End Sub
Private Sub PostPatInputBox_Change()
    Call ControlResponse(PostPatInputBox)
End Sub
Private Sub PostTestInputBox_Change()
    Call ControlResponse(PostTestInputBox)
End Sub
Private Sub EndOfBodyInputBox_Change()
    Call ControlResponse(EndOfBodyInputBox)
End Sub



'***************************************************
'NEXT 80 HANDLE CUSTOM ARGUMENTS
'***************************************************
Private Sub UserReqCheckBox1_Click()
    Call ControlResponse(UserReqCheckBox1)
End Sub
Private Sub UserReqComboBox1_Change()
    Call ControlResponse(UserReqComboBox1)
End Sub
Private Sub UserReqTextBox1_Change()
    Call ControlResponse(UserReqTextBox1)
End Sub
Private Sub UserReqButton1_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq1)
End Sub
Private Sub UserReqCheckBox2_Click()
    Call ControlResponse(UserReqCheckBox2)
End Sub
Private Sub UserReqComboBox2_Change()
    Call ControlResponse(UserReqComboBox2)
End Sub
Private Sub UserReqTextBox2_Change()
    Call ControlResponse(UserReqTextBox2)
End Sub
Private Sub UserReqButton2_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq2)
End Sub
Private Sub UserReqCheckBox3_Click()
    Call ControlResponse(UserReqCheckBox3)
End Sub
Private Sub UserReqComboBox3_Change()
    Call ControlResponse(UserReqComboBox3)
End Sub
Private Sub UserReqTextBox3_Change()
    Call ControlResponse(UserReqTextBox3)
End Sub
Private Sub UserReqButton3_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq3)
End Sub
Private Sub UserReqCheckBox4_Click()
    Call ControlResponse(UserReqCheckBox4)
End Sub
Private Sub UserReqComboBox4_Change()
    Call ControlResponse(UserReqComboBox4)
End Sub
Private Sub UserReqTextBox4_Change()
    Call ControlResponse(UserReqTextBox4)
End Sub
Private Sub UserReqButton4_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq4)
End Sub
Private Sub UserReqCheckBox5_Click()
    Call ControlResponse(UserReqCheckBox5)
End Sub
Private Sub UserReqComboBox5_Change()
    Call ControlResponse(UserReqComboBox5)
End Sub
Private Sub UserReqTextBox5_Change()
    Call ControlResponse(UserReqTextBox5)
End Sub
Private Sub UserReqButton5_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq5)
End Sub
Private Sub UserReqCheckBox6_Click()
    Call ControlResponse(UserReqCheckBox6)
End Sub
Private Sub UserReqComboBox6_Change()
    Call ControlResponse(UserReqComboBox6)
End Sub
Private Sub UserReqTextBox6_Change()
    Call ControlResponse(UserReqTextBox6)
End Sub
Private Sub UserReqButton6_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq6)
End Sub
Private Sub UserReqCheckBox7_Click()
    Call ControlResponse(UserReqCheckBox7)
End Sub
Private Sub UserReqComboBox7_Change()
    Call ControlResponse(UserReqComboBox7)
End Sub
Private Sub UserReqTextBox7_Change()
    Call ControlResponse(UserReqTextBox7)
End Sub
Private Sub UserReqButton7_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq7)
End Sub
Private Sub UserReqCheckBox8_Click()
    Call ControlResponse(UserReqCheckBox8)
End Sub
Private Sub UserReqComboBox8_Change()
    Call ControlResponse(UserReqComboBox8)
End Sub
Private Sub UserReqTextBox8_Change()
    Call ControlResponse(UserReqTextBox8)
End Sub
Private Sub UserReqButton8_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq8)
End Sub
Private Sub UserReqCheckBox9_Click()
    Call ControlResponse(UserReqCheckBox9)
End Sub
Private Sub UserReqComboBox9_Change()
    Call ControlResponse(UserReqComboBox9)
End Sub
Private Sub UserReqTextBox9_Change()
    Call ControlResponse(UserReqTextBox9)
End Sub
Private Sub UserReqButton9_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq9)
End Sub
Private Sub UserReqCheckBox10_Click()
    Call ControlResponse(UserReqCheckBox10)
End Sub
Private Sub UserReqComboBox10_Change()
    Call ControlResponse(UserReqComboBox10)
End Sub
Private Sub UserReqTextBox10_Change()
    Call ControlResponse(UserReqTextBox10)
End Sub
Private Sub UserReqButton10_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserReq10)
End Sub
Private Sub UserOptCheckBox1_Click()
    Call ControlResponse(UserOptCheckBox1)
End Sub
Private Sub UserOptComboBox1_Change()
    Call ControlResponse(UserOptComboBox1)
End Sub
Private Sub UserOptTextBox1_Change()
    Call ControlResponse(UserOptTextBox1)
End Sub
Private Sub UserOptButton1_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt1)
End Sub
Private Sub UserOptCheckBox2_Click()
    Call ControlResponse(UserOptCheckBox2)
End Sub
Private Sub UserOptComboBox2_Change()
    Call ControlResponse(UserOptComboBox2)
End Sub
Private Sub UserOptTextBox2_Change()
    Call ControlResponse(UserOptTextBox2)
End Sub
Private Sub UserOptButton2_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt2)
End Sub
Private Sub UserOptCheckBox3_Click()
    Call ControlResponse(UserOptCheckBox3)
End Sub
Private Sub UserOptComboBox3_Change()
    Call ControlResponse(UserOptComboBox3)
End Sub
Private Sub UserOptTextBox3_Change()
    Call ControlResponse(UserOptTextBox3)
End Sub
Private Sub UserOptButton3_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt3)
End Sub
Private Sub UserOptCheckBox4_Click()
    Call ControlResponse(UserOptCheckBox4)
End Sub
Private Sub UserOptComboBox4_Change()
    Call ControlResponse(UserOptComboBox4)
End Sub
Private Sub UserOptTextBox4_Change()
    Call ControlResponse(UserOptTextBox4)
End Sub
Private Sub UserOptButton4_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt4)
End Sub
Private Sub UserOptCheckBox5_Click()
    Call ControlResponse(UserOptCheckBox5)
End Sub
Private Sub UserOptComboBox5_Change()
    Call ControlResponse(UserOptComboBox5)
End Sub
Private Sub UserOptTextBox5_Change()
    Call ControlResponse(UserOptTextBox5)
End Sub
Private Sub UserOptButton5_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt5)
End Sub
Private Sub UserOptCheckBox6_Click()
    Call ControlResponse(UserOptCheckBox6)
End Sub
Private Sub UserOptComboBox6_Change()
    Call ControlResponse(UserOptComboBox6)
End Sub
Private Sub UserOptTextBox6_Change()
    Call ControlResponse(UserOptTextBox6)
End Sub
Private Sub UserOptButton6_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt6)
End Sub
Private Sub UserOptCheckBox7_Click()
    Call ControlResponse(UserOptCheckBox7)
End Sub
Private Sub UserOptComboBox7_Change()
    Call ControlResponse(UserOptComboBox7)
End Sub
Private Sub UserOptTextBox7_Change()
    Call ControlResponse(UserOptTextBox7)
End Sub
Private Sub UserOptButton7_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt7)
End Sub
Private Sub UserOptCheckBox8_Click()
    Call ControlResponse(UserOptCheckBox8)
End Sub
Private Sub UserOptComboBox8_Change()
    Call ControlResponse(UserOptComboBox8)
End Sub
Private Sub UserOptTextBox8_Change()
    Call ControlResponse(UserOptTextBox8)
End Sub
Private Sub UserOptButton8_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt8)
End Sub
Private Sub UserOptCheckBox9_Click()
    Call ControlResponse(UserOptCheckBox9)
End Sub
Private Sub UserOptComboBox9_Change()
    Call ControlResponse(UserOptComboBox9)
End Sub
Private Sub UserOptTextBox9_Change()
    Call ControlResponse(UserOptTextBox9)
End Sub
Private Sub UserOptButton9_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt9)
End Sub
Private Sub UserOptCheckBox10_Click()
    Call ControlResponse(UserOptCheckBox10)
End Sub
Private Sub UserOptComboBox10_Change()
    Call ControlResponse(UserOptComboBox10)
End Sub
Private Sub UserOptTextBox10_Change()
    Call ControlResponse(UserOptTextBox10)
End Sub
Private Sub UserOptButton10_Click()
    ApplyButton.enabled = tl_fs_ButtonClick(tl_tm_ParUserOpt10)
End Sub






'***************************************************
'NEXT 9 HANDLE START STATE CONDITIONS AND FLOAT PINS AND UTILITY PINS
'***************************************************
Private Sub Type3Button_Click()
    Call ControlResponse(Type3Button)
End Sub
Private Sub FloatEditButton_Click()
    Call ControlResponse(FloatEditButton, TL_C_FLOATSELFRMSTR, TL_C_FLOATSELAVLSTR, TL_C_FLOATSELSELSTR)
End Sub
Private Sub StartInitLObox_Change()
    Call ControlResponse(StartInitLOBox)
End Sub
Private Sub StartInitHIbox_Change()
    Call ControlResponse(StartInitHIBox)
End Sub
Private Sub StartInitZbox_Change()
    Call ControlResponse(StartInitZBox)
End Sub
Private Sub FloatPinsBox_Change()
    Call ControlResponse(FloatPinsBox)
End Sub
Private Sub Util1PinsBox_Change()
    Call ControlResponse(Util1PinsBox)
End Sub
Private Sub Util0PinsBox_Change()
    Call ControlResponse(Util0PinsBox)
End Sub
Private Sub UtilEditButton_Click()
    Call ControlResponse(UtilEditButton, TL_C_UTILSELFRMSTR, TL_C_UTILSELAVLSTR, TL_C_UTILSEL0STR, TL_C_UTILSEL1STR)
End Sub







Sub ControlResponse(CurObj As Object, Optional title1 As String, Optional title2 As String, Optional title3 As String, Optional title4 As String)
    For Each tl_tm_ParThisPar In AllPars
        If (tl_tm_ParThisPar.EquationButton Is CurObj) Or _
           (tl_tm_ParThisPar.SelectorButton Is CurObj) Or _
           (tl_tm_ParThisPar.FileFindButton Is CurObj) Or _
           (tl_tm_ParThisPar.DataBox Is CurObj) Or _
           (tl_tm_ParThisPar.CheckBox Is CurObj) Then
            Exit For
        End If
    Next
    If (tl_tm_ParThisPar.EquationButton Is CurObj) Then
        ApplyButton.enabled = tl_tm_ParThisPar.Equation(tl_tm_FormCtrl.EnableCtrl, tl_tm_FormCtrl.ChangedStatus)
    End If
    If (tl_tm_ParThisPar.SelectorButton Is CurObj) Then
        If CurObj.Name = Type3Button.Name Then
            ApplyButton.enabled = tl_tm_StartStateHiLoZ.ClickMade(tl_tm_FormCtrl.EnableCtrl, tl_tm_FormCtrl.ChangedStatus, TL_C_THREESELFRMSTR, TL_C_PINSELAVLSTR, TL_C_STATE0SELSELSTR, TL_C_STATE1SELSELSTR, TL_C_STATEZSELSELSTR)
        ElseIf CurObj.Name = UtilEditButton.Name Then
            ApplyButton.enabled = tl_tm_Utility10.ClickMade(tl_tm_FormCtrl.EnableCtrl, tl_tm_FormCtrl.ChangedStatus, TL_C_UTILSELFRMSTR, TL_C_UTILSELAVLSTR, TL_C_UTILSEL0STR, TL_C_UTILSEL1STR)
        Else
            ApplyButton.enabled = tl_tm_ParThisPar.ClickMade(tl_tm_FormCtrl.EnableCtrl, tl_tm_FormCtrl.ChangedStatus, title1, title2, title3)
        End If
    End If
    If (tl_tm_ParThisPar.FileFindButton Is CurObj) Then
        If tl_tm_ParThisPar.TestIsSinglePat = True Then
            Call tl_fs_PatSelectReplace(tl_tm_ParThisPar)
        Else
            Call tl_fs_PatSelectAppend(tl_tm_ParThisPar)
        End If
        ApplyButton.enabled = True
    End If
    If (tl_tm_ParThisPar.DataBox Is CurObj) Then
        ApplyButton.enabled = tl_tm_ParThisPar.ChangeMade(tl_tm_FormCtrl.EnableCtrl, tl_tm_FormCtrl.ChangedStatus, tl_tm_OpMode, RunButton)
    End If
    If (tl_tm_ParThisPar.CheckBox Is CurObj) Then
        If tl_tm_OpMode = TL_C_INTERACTIVE Then
            Dim ParStrPos As Integer
            ApplyButton.enabled = True
            With tl_tm_ParThisPar.CheckBoxArg
                'This is where we convert the information in the checkbox to a form that we will store in ParameterStr.
                If tl_tm_ParThisPar.CheckBox.Value = True Then
                    .ParameterStr = CStr(val(.ParameterStr) + 2 ^ (tl_tm_ParThisPar.CheckBoxBitPos))
                Else
                    .ParameterStr = CStr(val(.ParameterStr) - 2 ^ (tl_tm_ParThisPar.CheckBoxBitPos))
                End If
                If tl_tm_FormCtrl.EnableCtrl = True Then
                    tl_tm_FormCtrl.ChangedStatus = True
                    ApplyButton.enabled = True
                    .ValueChanged = True
                End If
                'to determine whether an EvalPrep should be performed on the data, see if an evalbox exists, and references to other boxes are nothing,
                '   or if the references are something whether the reference to the EvalBox is nothing.
                .ParameterValue = CStr(tl_dt_IEGetFormulaValue(.ParameterStr))
            End With
        End If
    End If
    If tl_tm_FormCtrl.EnableCtrl = True Then
        tl_tm_FormCtrl.ChangedStatus = True
        tl_tm_FormCtrl.ButtonEnabled = ApplyButton.enabled
    End If

End Sub










'***************************************************
'NEXT 1 HANDLES IE FORM VARIABLE INITIALIZATION
'***************************************************
Private Sub UserForm_Initialize()
    Dim FontName As String
    Dim FontSize As Integer
    Dim MemCnt As Integer
    Dim MemNum As Integer
    Dim lngX As Long
    Dim temp As String
    
    tl_tm_MultiItemMode = False

    If tl_tm_InstanceEditor.BpmuPages Or tl_tm_InstanceEditor.PpmuPages Then
        Me.ReqPages.PmuPage.Visible = True
    Else
        Me.ReqPages.PmuPage.Visible = False
    End If
    Me.OptPages.BpmuPage.Visible = tl_tm_InstanceEditor.BpmuPages
    Me.OptPages.PpmuPage.Visible = tl_tm_InstanceEditor.PpmuPages
    Me.ReqPages.DpsPage.Visible = tl_tm_InstanceEditor.DpsPages
    Me.OptPages.DpsPage.Visible = tl_tm_InstanceEditor.DpsPages
    Me.ReqPages.PatPage.Visible = tl_tm_InstanceEditor.FuncPage
    Me.OptPages.PatPage.Visible = tl_tm_InstanceEditor.CondPatPage
    
    Me.OptPages.LevTimPage.Visible = tl_tm_InstanceEditor.LevTimPage
    Me.OptPages.PinPage.Visible = tl_tm_InstanceEditor.PinPage
    Me.OptPages.PatFlagFuncPage.Visible = tl_tm_InstanceEditor.PatFuncPage
    Me.OptPages.InterposePage.Visible = tl_tm_InstanceEditor.InterposePage
    Me.SerializeMeasCheckBox.Visible = tl_tm_InstanceEditor.SerializeMeasureCheckBox
    Me.SerializeMeasFBox.Visible = tl_tm_InstanceEditor.SerializeMeasFBox
    Me.SerializeMeasFInputBox.Visible = tl_tm_InstanceEditor.SerializeMeasFInputBox
    Me.SerializeMeasFLabel.Visible = tl_tm_InstanceEditor.SerailizeMeasFLabel

    With Me
        .ReqPages.UserPage1.Visible = tl_tm_InstanceEditor.UserReqPage1
        .ReqPages.UserPage2.Visible = tl_tm_InstanceEditor.UserReqPage2
        .OptPages.UserPage1.Visible = tl_tm_InstanceEditor.UserOptPage1
        .OptPages.UserPage2.Visible = tl_tm_InstanceEditor.UserOptPage2
        Call tl_fs_SetPageObjectEnable(.ReqPages)
        Call tl_fs_SetPageObjectEnable(.OptPages)
    End With

    If (tl_tm_InstanceEditor.BpmuPages) Or (tl_tm_InstanceEditor.PpmuPages) Then
        If tl_tm_InstanceEditor.BpmuPages Then
            Me.ReqPages.PmuPage.Caption = "BPMU"
        Else
            Me.ReqPages.PmuPage.Caption = "PPMU"
        End If
    End If


    If TheBook.IsValid = True Then
        tl_tm_BookIsValid = True
    Else
        tl_tm_BookIsValid = False
    End If
    RunButton.enabled = tl_tm_BookIsValid
    
    If (tl_tm_SysInit = False) Then
        Call tl_fs_TemplateSystemInit(tl_tm_SysInit)
    End If
    With JobData
        Call tl_fs_TemplateCatSelStrings(.AvailDcCat, .AvailDcSel, _
            .AvailAcCat, .AvailAcSel, .AvailTimeSetAll, _
            .AvailTimeSetExtended, .AvailEdgeSet, .AvailLevels)
        Call tl_fs_TemplateJobDataPinlistStrings(JobData)
        Call tl_fs_TemplateSpecsPatsStrings(.AllSpecs, .AllSpecEntries, .AllPatNames, .AllPatSetNames, .AllPatGrpNames)
        'Get list of Overlay
        Call tl_fs_TemplateOverlayString(.AvailOverlay)
    End With
    tl_tm_FormCtrl.ChangedStatus = False
    tl_tm_FormCtrl.EnableCtrl = False
    
    Call ParameterSetup
    For Each tl_tm_ParThisPar In AllPars
        Call tl_fs_SetupFormObject(tl_tm_ParThisPar)
    Next
    
    'Initialize the IE form.
    ApplyButton.enabled = False
    RequiredFrame.Caption = TL_C_REQSTR
    OptionalFrame.Caption = TL_C_OPTSTR
    
    'update the caption as well
    Call tl_fs_SetCaption(Me, tl_tm_InstanceEditor.Caption)
    
    'provide text for buttons
    ApplyButton.Caption = TL_C_ApplySTR
    ExitButton.Caption = TL_C_okSTR
    ExitButton.default = True
    CancelButton.Caption = TL_C_CancelSTR
    CancelButton.Cancel = True
    HelpButton.Caption = TL_C_HelpSTR
    GroupCountBox.Caption = TL_C_IDXLIMSTR
    GroupNumLabel.Caption = TL_C_MEMIDXSTR
    
    'provide text for IPF headers
    IpfNameLabel.Caption = TL_C_NameSTR
    IpfValueLabel.Caption = TL_C_ValuSTR
    IpfParStrLabel.Caption = TL_C_ParaSTR

    tl_tm_OpMode = TL_C_BACKGROUND
    Call ReadSheet_SetForm
    
    If Me.OptPages.LevTimPage.enabled Then
        If (JobData.AvailDcCat = TL_C_EMPTYSTR) And (tl_tm_ParDcCat.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParDcCat)
        End If
        If (JobData.AvailDcSel = TL_C_EMPTYSTR) And (tl_tm_ParDcSel.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParDcSel)
        End If
        If (JobData.AvailAcCat = TL_C_EMPTYSTR) And (tl_tm_ParAcCat.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParAcCat)
        End If
        If (JobData.AvailAcSel = TL_C_EMPTYSTR) And (tl_tm_ParAcSel.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParAcSel)
        End If
        If (JobData.AvailTimeSetAll = TL_C_EMPTYSTR) And (tl_tm_ParTimeSet.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParTimeSet)
        End If
        If (JobData.AvailEdgeSet = TL_C_EMPTYSTR) And (tl_tm_ParEdgeSet.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParEdgeSet)
        End If
        If (JobData.AvailLevels = TL_C_EMPTYSTR) And (tl_tm_ParLevels.ParameterStr = TL_C_EMPTYSTR) Then
            Call tl_fs_MakeGray(tl_tm_ParLevels)
        End If
    End If
    
    Call tl_fs_MemberCount(MemCnt, MemNum)
    GroupCountBox.Caption = TL_C_MAXIDXSTR & TL_C_BLANKCHR & CStr(MemCnt)
    GroupNumBox.Text = CStr(MemNum)
    Call ControlButtons
    If tl_tm_MultiItemMode Then
        'use the first data item as the items values
        
        temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsHiLimSpec.ParameterItemsStr, 0)
        tl_tm_ParDpsHiLimSpec.ParameterStr = tl_tm_GetItemForm(temp)
        tl_tm_ParDpsHiLimSpec.DataBox.Text = tl_tm_ParDpsHiLimSpec.ParameterStr
        
        temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsLoLimSpec.ParameterItemsStr, 0)
        tl_tm_ParDpsLoLimSpec.ParameterStr = tl_tm_GetItemForm(temp)
        tl_tm_ParDpsLoLimSpec.DataBox.Text = tl_tm_ParDpsLoLimSpec.ParameterStr
        
        temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsSettlingTime.ParameterItemsStr, 0)
        tl_tm_ParDpsSettlingTime.ParameterStr = tl_tm_GetItemForm(temp)
        tl_tm_ParDpsSettlingTime.DataBox.Text = tl_tm_ParDpsSettlingTime.ParameterStr
        
        temp = tl_tm_GetItemFormAndVal(tl_tm_ParDpsLimits.ParameterItemsStr, 0)
        tl_tm_ParDpsLimits.ParameterStr = tl_tm_GetItemForm(temp)
        Call tl_fs_SetCheckBoxes(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, _
            tl_tm_ParDpsHiLimSpec, tl_tm_ParDpsLoLimSpec)
        
    End If
    
    tl_tm_OpMode = TL_C_INTERACTIVE
    Call ValData(TL_C_VALDATAMODENOSTOP)
    
    FontName = tl_tm_MsgStr(TL_TM_FONTNAME)
    FontSize = CInt(val(tl_tm_MsgStr(TL_TM_FONTSIZE)))
    
    For Each tl_tm_ParThisPar In AllPars
        If Not (tl_tm_ParThisPar.EvalBox Is Nothing) Then
            tl_tm_ParThisPar.EvalBox.Font.Name = FontName
            tl_tm_ParThisPar.EvalBox.Font.Size = FontSize
        End If
    Next
    
    If tl_tm_FocusArg <> 0 Then
        ' set the focus on the element needed...
        Call tl_fs_InitFocus(tl_tm_FocusArg, AllPars)
    End If
    
    Me.HelpContextID = tl_tm_InstanceEditor.HelpValue

    '------------------------------------------------------------------------------------
    'Eee-JOB DCTestScenarioÉCÉìÉXÉ^ÉìÉXÉGÉfÉBÉ^ópèàóùïœçX
    'DC Test ScenarioêÍópÉGÉfÉBÉ^ÇÃÇΩÇﬂÇÃÉRÉìÉgÉçÅ[ÉãäeéÌÇÃê›íËÇí«â¡
    IECustomForDCTestScenario
    '------------------------------------------------------------------------------------

End Sub

'***************************************************
'NEXT 1 HANDLES FORM DISPLAY INITIALIZATION
'***************************************************
Private Sub ReadSheet_SetForm()
    Dim temp As String
    
    Call tl_fs_GetTemplateArgs

    'Pattern Conditioning - CondPatPage,
    If tl_tm_InstanceEditor.CondPatPage Then
        Call tl_fs_SetWaitFlags(tl_tm_ParFlags, A1CheckBox, A0CheckBox, AXCheckBox, _
            B1CheckBox, B0CheckBox, BXCheckBox, _
            C1CheckBox, C0CheckBox, CXCheckBox, _
            D1CheckBox, D0CheckBox, DXCheckBox)
    End If
    If tl_tm_InstanceEditor.PatFuncPage Then
        Call tl_fs_SetWaitFlags(tl_tm_ParPatFlags, PatA1CheckBox, PatA0CheckBox, PatAXCheckBox, _
            PatB1CheckBox, PatB0CheckBox, PatBXCheckBox, _
            PatC1CheckBox, PatC0CheckBox, PatCXCheckBox, _
            PatD1CheckBox, PatD0CheckBox, PatDXCheckBox)
    End If
    
    'Ppmu - HiLoLimValid,
    If tl_tm_InstanceEditor.PpmuPages Then
        tl_tm_ParPpmuLimits.ParameterStr = tl_tm_ParPpmuLimits.GetParStr
        Call tl_fs_SetCheckBoxes(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParPpmuLimits, _
            tl_tm_ParPpmuHiLimSpec, tl_tm_ParPpmuLoLimSpec)
    End If
    'Bpmu - HiLoLimValid,
    If tl_tm_InstanceEditor.BpmuPages Then
        tl_tm_ParBpmuLimits.ParameterStr = tl_tm_ParBpmuLimits.GetParStr
        Call tl_fs_SetCheckBoxes(PmuHiLimitCheckBox, PmuLoLimitCheckBox, tl_tm_ParBpmuLimits, _
            tl_tm_ParBpmuHiLimSpec, tl_tm_ParBpmuLoLimSpec)
    End If
    'Dps - HiLoLimValid,
    If tl_tm_InstanceEditor.DpsPages Then
        tl_tm_ParDpsLimits.ParameterStr = tl_tm_ParDpsLimits.GetParStr
        Call tl_fs_SetCheckBoxes(DpsHiLimitCheckBox, DpsLoLimitCheckBox, tl_tm_ParDpsLimits, _
            tl_tm_ParDpsHiLimSpec, tl_tm_ParDpsLoLimSpec)
        tl_tm_ParSerializeMeas.ParameterStr = tl_tm_ParSerializeMeas.GetParStr
        If (tl_tm_InstanceEditor.SerializeMeasureCheckBox = True) Then
            Call tl_fs_SetSerializeMeasCheckBox(SerializeMeasCheckBox, tl_tm_ParSerializeMeas)
        End If
        
        
        
    End If
    
    ' now enable proper registration of data changing
    tl_tm_FormCtrl.EnableCtrl = True

End Sub



'***************************************************
'NEXT 1 HANDLES SAVING FORM DATA TO INSTANCE SHEET
'***************************************************
Private Sub SaveFormData()
    
    'save parameters data into the instance sheet
    For Each tl_tm_ParThisPar In AllPars
        If tl_tm_ParThisPar.ValueChanged = True Then
            tl_tm_ParThisPar.SaveData
        End If
    Next
    
End Sub


'***************************************************
'NEXT 1 HANDLES VALIDATION OF PARAMETERS
'***************************************************
Private Function ValData(VDCint As Integer) As Boolean
'ValData is the routine for the IE to Validate the Data.
'this subroutine is used, prior to saving data to the instance sheet, to
'determine whether the data to be saved is proper, valid, and copacetic.
'   It will be run in a particular 'mode', and in some modes, the
'   user will interact, and these user interactions may change the mode.
'   The modes of calling this routine
'   are:
'   TL_C_VALDATAMODENORMAL  -   Fix the current parameter being evaluated.
'   TL_C_VALDATAMODENOSTOP  -   Do not stop to fix any parameters.
'
    Dim temp As String
    Dim TestResult As Integer
    Dim intX As Integer

    If (VDCint <> TL_C_VALDATAMODENORMAL) And (VDCint <> TL_C_VALDATAMODENOSTOP) Then
        VDCint = TL_C_VALDATAMODENORMAL
    End If

    '   ValData will call a routine specific to each template type, to
    '   perform specific validation tasks.  That routine is .ValidateParameters,
    '   and there is one of those within each Template module.
    '   That has modes to run in as well.  If a mode of '0' is specified for
    '   input there, it is assumed that the mode is TL_C_VALDATAMODEJOBVAL.
    '   The modes of .ValidateParameters that are of interest here:
    '   TL_C_VALDATAMODENORMAL  -   Instance Editor mode; Fix the current parameter being evaluated.
    '   TL_C_VALDATAMODENOSTOP  -   Instance Editor mode; Do not stop to fix any parameters.
    '   It can come back with different modes, such as:
    '   TL_C_VALDATAMODENOFIX   -   Instance Editor mode; Error found, that specific one was not fixed.
    '   TL_C_VALDATAMODEFIXNONE -   Instance Editor mode; Error(s) found, none were fixed.

    'call the ValidateParameters routine for the Teradyne template, or Custom template
    temp = tl_dt_IEGetTemplateName
    intX = InStr(temp, "!")
    If intX <> 0 Then
        temp = temp & ".ValidateParameters"
        On Error GoTo RunErr
        'run the proper ValidateParameters routine in the specified mode, and get
        '   back a result, and the mode which the user may have changed!
        TestResult = Application.Run(temp, VDCint)
        On Error GoTo 0
        If TestResult = TL_SUCCESS Then
            ValData = True
        Else
            ValData = False
        End If
    Else
        'denote an error
        Call tl_ErrorLogMessage("InstanceEditor: tl_dt_IEGetTemplateName " & TL_C_ERRORSTR & " : " & temp)
        Call tl_ErrorReport
        Exit Function
    End If
    
    'AT THIS POINT, IF VALDATA=FALSE, THAT MEANS THAT THE USER HAS CHOSEN NOT TO 'FIX' THE
    'PROBLEMATIC PARAMETERS IN NEED OF ATTENTION.
    'PRESENT A MSGBOX ASKING THE USER to acknowledge this, if in 'NoFix' or 'FixNone' Mode
    If (ValData = False) And _
        ((VDCint = TL_C_VALDATAMODENOFIX) Or (VDCint = TL_C_VALDATAMODEFIXNONE)) Then
        Call MsgBox(tl_tm_MsgStr(TL_TM_STR_PROB), vbOKOnly, tl_tm_MsgStr(TL_TM_STR_YESNONOTOALL))
    End If
    Exit Function
    
RunErr:
    On Error GoTo 0
    'denote error
    Call tl_ErrorLogMessage("InstanceEditor: ValidateParameters " & TL_C_ERRORSTR)
    Call tl_ErrorReport
End Function

Private Sub IECustomForDCTestScenario()
'ì‡óe:
'   Eee-JOB DC Test ScenarioÇÃÇΩÇﬂÇÃÉGÉfÉBÉ^ÉtÉHÅ[ÉÄêÆå`
'   åªíiäKÇ≈ïsóvÇ»ÉRÉìÉgÉçÅ[ÉãÇÃîÒï\é¶Ç∆ÉTÉCÉYÅAà íuí≤êÆÇÃê›íË
'
'ÉpÉâÉÅÅ[É^:
'
'ñﬂÇËíl:
'
'íçà”éñçÄ:
'
    With Me
        .height = 237.75
        .width = 412.25
        .RequiredFrame.Visible = False
        With .OptionalFrame
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .Caption = "DC Test Scenario Execute Arguments"
            .width = 396
            .height = 120
            .Top = 6
            .Left = 6
        End With
        With .OptPages
            .width = 378
            .height = 90
            .Left = 6
            .Top = 12
        End With
        .RunButton.Visible = False
        .CopyButton.Visible = False
        .PasteButton.Visible = False
        .DeleteButton.Visible = False
        .AddButton.Visible = False
        .HelpButton.Visible = False
        .GroupCountBox.Visible = False
        .GroupNumLabel.Visible = False
        .GroupNumSpinButton.Visible = False
        .GroupCountBox.Visible = False
        .GroupNumBox.Visible = False
        .PrePatFBox.Visible = False
        .PrePatFLabel.Visible = False
        .PrePatInputBox.Visible = False
        .PreTestBox.Visible = False
        .PreTestFLabel.Visible = False
        .PreTestInputBox.Visible = False
        .PostPatFBox.Visible = False
        .PostPatFLabel.Visible = False
        .PostPatInputBox.Visible = False
        .PostTestBox.Visible = False
        .PostTestFLabel.Visible = False
        .PostTestInputBox.Visible = False
        With .IpfNameLabel
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .AutoSize = True
            .Left = 30
            .Top = 6
        End With
        With .IpfValueLabel
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .AutoSize = True
            .Left = 114
            .Top = 6
        End With
        With .IpfParStrLabel
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .AutoSize = True
            .Left = 240
            .Top = 6
        End With
        With .StartOfBodyFLabel
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .AutoSize = True
            .Left = 12
            .Top = 24
        End With
        With .StartOfBodyBox
            .width = 108
            .Left = 72
            .Top = 20
        End With
        With .StartOfBodyInputBox
            .width = 174
            .Left = 186
            .Top = 20
        End With
        With .EndOfBodyFLabel
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .AutoSize = True
            .Left = 16.5
            .Top = 48
        End With
        With .EndOfBodyBox
            .width = 108
            .Left = 72
            .Top = 46
        End With
        With .EndOfBodyInputBox
            .width = 174
            .Left = 186
            .Top = 46
        End With
        With .CommentsLabel
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .AutoSize = True
            .Left = 18
            .Top = 150
        End With
        With .CommentsTextBox
            .width = 330
            .height = 33
            .Top = 138
            .Left = 72
        End With
        With .ExitButton
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .Top = 186
            .Left = 18
        End With
        With .CancelButton
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .Top = 186
            .Left = 90
        End With
        With .ApplyButton
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .Top = 186
            .Left = 162
        End With
        With .EeeJOBLabel
            .Visible = True
            .Font.Name = "Tahoma"
            .Font.Size = 8
'            .Caption = APPNAME_PREFIX & APPLICATION_NAME & " " & REV_PREFIX & XLibAddInInfo.GetRevisionNumber
            .Caption = ""
            .AutoSize = True
            .width = 150
            .Left = 250
            .Top = 194
        End With
    End With
End Sub
