Attribute VB_Name = "CodePE8API"
Option Explicit
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS
' FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER
' IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

' Implicitly declared PE constants (used by code but declared by the graph control)
'Global Const SGPM_BAR& = 10
'Global Const PEABT_THIN_LINE& = 2
'Global Const REINITIALIZE_RESETIMAGE& = 0
'Global Const UNDO_ZOOM& = 19
'Global Const PEANF_EXP_NOTATION& = 1
'Global Const PEANF_EXP_NOTATION_3X& = 2
'Global Const PERE_GDIPLUS& = 2
'Global Const SGPM_POINT& = 1
'Global Const PEPGS_NONE& = 0
'Global Const PEMPS_MEDIUM_LARGE& = 5
'Global Const PELT_THIN_SOLID& = 0
'Global Const PELS_1_LINE_INSIDE_OVERLAP& = 4
'Global Const REVERT_TO_DEFAULTS& = 20

' Expicitly declared PE constants
'Global Const FIRST_DEFAULT_TAB& = 0

'Global Const PEDP_ENABLED& = 0
'Global Const PEDP_DISABLED& = 1
'Global Const PEDP_INSIDE_TOP& = 2

'Global Const PETLT_12HR_AM_PM& = 0
'Global Const PETLT_12HR_NO_AM_PM& = 1
'Global Const PETLT_24HR& = 2

'Global Const PEDLT_3_CHAR& = 0
'Global Const PEDLT_1_CHAR& = 1
'Global Const PEDLT_NO_DAY_PROMPT& = 2
'Global Const PEDLT_NO_DAY_NUMBER& = 3

'Global Const PEMLT_3_CHAR& = 0
'Global Const PEMLT_1_CHAR& = 1
'Global Const PEMLT_NO_MONTH_PROMPT& = 2

'Global Const PEHS_HORIZONTAL& = 0            '/* ----- */
'Global Const PEHS_VERTICAL& = 1              '/* ||||| */
'Global Const PEHS_FDIAGONAL& = 2             '/* \\\\\ */
'Global Const PEHS_BDIAGONAL& = 3             '/* ///// */
'Global Const PEHS_CROSS& = 4                 '/* +++++ */
'Global Const PEHS_DIAGCROSS& = 5             '/* xxxxx */

'Global Const PEGS_NO_GRADIENT& = 0
'Global Const PEGS_VERTICAL& = 1
'Global Const PEGS_HORIZONTAL& = 2

'Global Const PEBS_NO_BMP& = 0
'Global Const PEBS_STRETCHBLT& = 1
'Global Const PEBS_TILED_BITBLT& = 2
'Global Const PEBS_BITBLT_TOP_LEFT& = 3
'Global Const PEBS_BITBLT_TOP_CENTER& = 4
'Global Const PEBS_BITBLT_TOP_RIGHT& = 5
'Global Const PEBS_BITBLT_BOTTOM_LEFT& = 6
'Global Const PEBS_BITBLT_BOTTOM_CENTER& = 7
'Global Const PEBS_BITBLT_BOTTOM_RIGHT& = 8
'Global Const PEBS_BITBLT_CENTER& = 9

'Global Const PEQS_NO_STYLE& = 0
'Global Const PEQS_LIGHT_INSET& = 1
'Global Const PEQS_LIGHT_SHADOW& = 2
'Global Const PEQS_LIGHT_LINE& = 3
'Global Const PEQS_LIGHT_NO_BORDER& = 4
'Global Const PEQS_MEDIUM_INSET& = 5
'Global Const PEQS_MEDIUM_SHADOW& = 6
'Global Const PEQS_MEDIUM_LINE& = 7
'Global Const PEQS_MEDIUM_NO_BORDER& = 8
'Global Const PEQS_DARK_INSET& = 9
'Global Const PEQS_DARK_SHADOW& = 10
'Global Const PEQS_DARK_LINE& = 11
'Global Const PEQS_DARK_NO_BORDER& = 12

 '// PEP_dwDESKCOLOR must be set to 1(PEBG_TRANSPARENT) to enable the below properties //
'Global Const PEP_dwDESKGRADIENTSTART& = 2687
'Global Const PEP_dwDESKGRADIENTEND& = 2688
'Global Const PEP_nDESKGRADIENTSTYLE& = 2689
'Global Const PEP_szDESKBMPFILENAME& = 2690
'Global Const PEP_nDESKBMPSTYLE& = 2691

'// PEP_dwGRAPHBACKCOLOR must be set to 1 (PEBG_TRANSPARENT) to enable the below properties.
'Global Const PEP_dwGRAPHGRADIENTSTART& = 2692
'Global Const PEP_dwGRAPHGRADIENTEND& = 2693
'Global Const PEP_nGRAPHGRADIENTSTYLE& = 2694
'Global Const PEP_szGRAPHBMPFILENAME& = 2695
'Global Const PEP_nGRAPHBMPSTYLE& = 2696

'// PEP_dwTABLEBACKCOLOR must be set to 1 (PEBG_TRANSPARENT) to enable the below properties.
'Global Const PEP_dwTABLEGRADIENTSTART& = 2697
'Global Const PEP_dwTABLEGRADIENTEND& = 2698
'Global Const PEP_nTABLEGRADIENTSTYLE& = 2699
'Global Const PEP_szTABLEBMPFILENAME& = 2700
'Global Const PEP_nTABLEBMPSTYLE& = 2701
'Global Const PEBG_TRANSPARENT& = 1
'Global Const PEP_bPNGISTRANSPARENT& = 2683
'Global Const PEP_dwPNGTRANSPARENTCOLOR& = 2684
'Global Const PEP_bPNGISINTERLACED& = 2685
'Global Const PEP_nJPGQUALITY& = 2686
'Global Const PEP_bDISABLE3DSHADOW& = 3927
'Global Const PEP_nDROPSHADOWOFFSETX& = 2679
'Global Const PEP_nDROPSHADOWOFFSETY& = 2680
'Global Const PEP_nDROPSHADOWSTEPS& = 2681
'Global Const PEP_nDROPSHADOWWIDTH& = 2682
'Global Const PEP_nHIDEINTERSECTINGTEXT& = 2678
'Global Const PEP_bSTOP& = 2677
'Global Const PEP_nBITMAPGRADIENTMENU& = 2702
'Global Const PEP_bBITMAPGRADIENTMODE& = 2703
'Global Const PEP_nLONGXAXISTICKMENU& = 2674
'Global Const PEP_nLONGYAXISTICKMENU& = 2673
'Global Const PEP_nQUICKSTYLE& = 2672
'Global Const PEP_nQUICKSTYLEMENU& = 2671
'Global Const PEP_nVIEWINGSTYLEMENU& = 2640
'Global Const PEP_nFONTSIZEMENU& = 2641
'Global Const PEP_nDATAPRECISIONMENU& = 2642
'Global Const PEP_nDATASHADOWMENU& = 2643
'Global Const PEP_bSEPARATORMENU& = 2654
'Global Const PEP_nMAXIMIZEMENU& = 2655
'Global Const PEP_nCUSTOMIZEDIALOGMENU& = 2656
'Global Const PEP_nEXPORTDIALOGMENU& = 2657
'Global Const PEP_nHELPMENU& = 2658

'Global Const PEP_nBORDERTYPEMENU& = 2659
'Global Const PEP_nSHOWLEGENDMENU& = 2660
'Global Const PEP_nLEGENDLOCATIONMENU& = 2661
'Global Const PEP_nSHOWTABLEANNOTATIONSMENU& = 2662
'Global Const PEP_nMULTIAXISSTYLEMENU& = 2663
'Global Const PEP_nFIXEDFONTMENU& = 2664

'Global Const PEP_bSHOWALLTABLEANNOTATIONS& = 2665
'Global Const PEP_bSHOWLEGEND& = 2666

'Global Const PEP_naCUSTOMMENU& = 2667
'Global Const PEP_naCUSTOMMENUSTATE& = 2668
'Global Const PEP_naCUSTOMMENULOCATION& = 2669
'Global Const PEP_szaCUSTOMMENUTEXT& = 2670
'Global Const PEP_nLASTMENUINDEX& = 2675
'Global Const PEP_nLASTSUBMENUINDEX& = 2676

'Global Const PEP_nGRIDLINEMENU& = 3164
'Global Const PEP_nPLOTMETHODMENU& = 3165
'Global Const PEP_nGRIDINFRONTMENU& = 3166
'Global Const PEP_nTREATCOMPARISONSMENU& = 3167
'Global Const PEP_nMARKDATAPOINTSMENU& = 3168
'Global Const PEP_nSHOWANNOTATIONSMENU& = 3169
'Global Const PEP_nUNDOZOOMMENU& = 3170

'Global Const PEP_nGRAPHPLUSTABLEMENU& = 3430
'Global Const PEP_nFORCEVERTPOINTSMENU& = 3431
'Global Const PEP_nTABLEWHATMENU& = 3432

'Global Const PEP_nINCLUDEDATALABELSMENU& = 3696
'Global Const PEP_fZOOMMINTX& = 3697
'Global Const PEP_fZOOMMAXTX& = 3698

'Global Const PEP_nSHOWBOUNDINGBOXMENU& = 4058
'Global Const PEP_nROTATIONMENU& = 4059
'Global Const PEP_nCONTOURMENU& = 4060

'Global Const PEP_nPERCENTORVALUEMENU& = 3925
'Global Const PEP_nGROUPPERCENTMENU& = 3926
                                        
'Global Const PEP_nIMAGEADJUSTLEFT& = 2982
'Global Const PEP_nIMAGEADJUSTRIGHT& = 2983
'Global Const PEP_nIMAGEADJUSTTOP& = 2984
'Global Const PEP_nIMAGEADJUSTBOTTOM& = 2985

'Global Const PEP_bMODALDIALOGS& = 2978
'Global Const PEP_bMODELESSONTOP& = 2979
'Global Const PEP_bMODELESSAUTOCLOSE& = 2980
'Global Const PEP_nDIALOGRESULT& = 2981

'Global Const PEP_fPOINTPADDING& = 3427
'Global Const PEP_fPOINTPADDINGAREA& = 3428
'Global Const PEP_fPOINTPADDINGBAR& = 3429

'Global Const PEP_bCUSTOMGRIDNUMBERSY& = 3160
'Global Const PEP_bCUSTOMGRIDNUMBERSRY& = 3161
'Global Const PEP_bCUSTOMGRIDNUMBERSX& = 3163
'Global Const PEP_bCUSTOMGRIDNUMBERSTX& = 3695
'Global Const PEP_structCUSTOMGRIDNUMBERS& = 3162
'Global Const PEP_bCUSTOMGRIDNUMBERSZ& = 4055

'Global Const PEP_nTICKSTYLE& = 3158
'Global Const PEP_dwTICKCOLOR& = 3159

'Global Const PEP_naPOINTTYPES& = 3156
'Global Const PEP_naSUBSETFORPOINTTYPES& = 3157

'Global Const PEP_naSUBSETFORPOINTCOLORS& = 3155
'Global Const PEP_nZOOMSTYLE& = 3154

'Global Const PEP_nSHOWPIELABELS& = 3921
'Global Const PEP_bSHOWPIELEGEND& = 3922
'Global Const PEP_nSLICEHATCHING& = 3923
'Global Const PEP_nSLICESTARTLOCATION& = 3924
'Global Const PEP_nMULTIAXISSEPARATORSIZE& = 3153
'Global Const PEP_nCURSORPROMPTLOCATION& = 3152
'Global Const PEP_szaMULTIAXISTITLES& = 3150
'Global Const PEP_nLEGENDSTYLE& = 2975
'Global Const PEP_bNOSMARTTABLEPLACEMENT& = 2976
'Global Const PEP_nSMARTLEGENDTHRESHOLD& = 3148
'Global Const PEP_nMULTIAXISSTYLE& = 3149
'Global Const PEP_bFLOATINGBARS& = 3151
'Global Const PEP_bSIMPLELINELEGEND& = 2973
'Global Const PEP_bSIMPLEPOINTLEGEND& = 2974
'Global Const PEP_bDISABLESORTPLOTMETHODS& = 3147
'Global Const PEP_nWORKINGTABLE& = 2977
'Global Const PEP_nTAROWS& = 2951
'Global Const PEP_nTACOLUMNS& = 2952
'Global Const PEP_naTATYPE& = 2953
'Global Const PEP_szaTATEXT& = 2954
'Global Const PEP_dwaTACOLOR& = 2955
'Global Const PEP_naTAHOTSPOT& = 2956
'Global Const PEP_nTAHEADERROWS& = 2957
'Global Const PEP_bTAHEADERCOLUMN& = 2958
'Global Const PEP_naTACOLUMNWIDTH& = 2959
'Global Const PEP_nTAHEADERORIENTATION& = 2960
'Global Const PEP_nTALOCATION& = 2961
'Global Const PEP_nTABORDER& = 2962
'Global Const PEP_dwTABACKCOLOR& = 2963
'Global Const PEP_dwTAFORECOLOR& = 2964
'Global Const PEP_nTATEXTSIZE& = 2965
'Global Const PEP_nTAAXISLOCATION& = 2966
'Global Const PEP_bSHOWTABLEANNOTATION& = 2968
'Global Const PEP_naTAJUSTIFICATION& = 2969
'Global Const PEP_szTAFONT& = 2970
'Global Const PEP_szaTAFONTS& = 2971

'Global Const PEP_nDELIMITER& = 2950
'Global Const PEP_bDISABLESYMBOLFIX& = 2972

'Global Const PEP_nAXISSIZEY& = 3143
'Global Const PEP_nAXISLOCATIONY& = 3144
'Global Const PEP_nAXISSIZERY& = 3145
'Global Const PEP_nAXISLOCATIONRY& = 3146

'Global Const PEP_fFONTSIZEMSCNTL& = 2945
'Global Const PEP_fFONTSIZEMBCNTL& = 2946
'Global Const PEP_fFONTSIZEGNCNTL& = 2947
'Global Const PEP_fFONTSIZECPCNTL& = 2948
'Global Const PEP_fFONTSIZEALCNTL& = 2949

'Global Const PEP_bFIXEDLINETHICKNESS& = 3140
'Global Const PEP_bFIXEDSPMWIDTH& = 3141
'Global Const PEP_fDASHLINETHICKNESS& = 3142

'Global Const PEP_naHORZLINEANNOTHOTSPOT& = 3138
'Global Const PEP_naVERTLINEANNOTHOTSPOT& = 3139

'Global Const PEP_nYEARMONTHDAYPROMPT& = 3133
'Global Const PEP_nTIMELABELTYPE& = 3134
'Global Const PEP_nDAYLABELTYPE& = 3135
'Global Const PEP_nMONTHLABELTYPE& = 3136
'Global Const PEP_nYEARLABELTYPE& = 3137

'Global Const PEP_dwaAPPENDPOINTCOLORS& = 3132
'Global Const PEP_bDISABLECLIPPING& = 2944
'Global Const PEP_faWORKINGAXESPROPORTIONS& = 3131
'Global Const PEP_nBORDERTYPES& = 2943

'Global Const PEP_bDATETIMESHOWSECONDS& = 3129

'Global Const PEP_structSPRINGDAYLIGHT& = 3127
'Global Const PEP_structFALLDAYLIGHT& = 3128

'Global Const PEP_structEXTRAAXISX& = 3693
'Global Const PEP_structEXTRAAXISTX& = 3694

'Global Const PEP_bTRIANGLEANNOTATIONADJ& = 3126
'Global Const PEP_fGRIDASPECT& = 3124
'Global Const PEP_faGRIDHOTSPOTVALUE& = 3123

'Global Const PEP_bVGNAXISLABELLOCATION& = 3121
'Global Const PEP_bALLOWGRIDNUMBERHOTSPOTSY& = 3122
'Global Const PEP_bALLOWGRIDNUMBERHOTSPOTSX& = 3692

'Global Const PEP_fLEFTEDGESPACING& = 3117
'Global Const PEP_fRIGHTEDGESPACING& = 3118
'Global Const PEP_fAXISNUMBERSPACING& = 3119

'Global Const PEP_fPOLARTICKTHRESHOLD& = 3804
'Global Const PEP_fPOLARLINETHRESHOLD& = 3805
'Global Const PEP_fPOLAR30DEGTHRESHOLD& = 3806

'Global Const PEP_bOLDSCALINGLOGIC& = 2942

'Global Const PEP_bSHADEDPOLYGONBORDERS& = 4056

'Global Const PEP_dwBARBORDERCOLOR& = 3116

'Global Const PEP_dwHATCHBACKCOLOR& = 2941
'Global Const PEP_naSUBSETHATCH& = 2940
'Global Const PEP_naPOINTHATCH& = 3114

'Global Const PEP_bYAXISVERTGRIDNUMBERS& = 3113

'Global Const PEP_bDAYLIGHTSAVINGS& = 3112

'Global Const PEP_bFIXEDFONTS& = 2938
'Global Const PEP_hSIZENSCURSOR& = 2939

'Global Const PEP_bCONTOURSTYLELEGEND& = 3690
'Global Const PEP_szaCONTOURLABELS& = 3691

'Global Const PEP_naPLOTTINGMETHODS& = 3103
'Global Const PEP_nSPEEDBOOST& = 3104

'Global Const PEP_nSHOWTICKMARKY& = 3106
'Global Const PEP_nSHOWTICKMARKRY& = 3107
'Global Const PEP_nSHOWTICKMARKX& = 3108
'Global Const PEP_nOHLCMINWIDTH& = 3109
'Global Const PEP_nMULTIAXESSIZING& = 3111

'Global Const PEP_nSHOWTICKMARKTX& = 3689

'Global Const PEP_bFLOATINGSTACKEDBARS& = 3424
'Global Const PEP_nSCROLLINGRANGE& = 3425
'Global Const PEP_nSCROLLINGFACTOR& = 3426

'// FUNCTION / PROPERTY HELPER CODES //
'Global Const PECONTROL_GRAPH& = 300
'Global Const PECONTROL_PIE& = 302
'Global Const PECONTROL_SGRAPH& = 304
'Global Const PECONTROL_PGRAPH& = 308
'Global Const PECONTROL_3D& = 312

'Global Const PESTA_CENTER& = 0
'Global Const PESTA_LEFT& = 1
'Global Const PESTA_RIGHT& = 2

'Global Const PEDO_DRIVERDEFAULT& = 0
'Global Const PEDO_LANDSCAPE& = 1
'Global Const PEDO_PORTRAIT& = 2

'Global Const PEVS_COLOR& = 0
'Global Const PEVS_MONO& = 1
'Global Const PEVS_MONOWITHSYMBOLS& = 2

Global Const PEFS_LARGE& = 0
Global Const PEFS_MEDIUM& = 1
Global Const PEFS_SMALL& = 2

'Global Const PEVB_NONE& = 0
'Global Const PEVB_TOP& = 1
'Global Const PEVB_BOTTOM& = 2
'Global Const PEVB_TOPANDBOTTOM& = 3

Global Const PEAC_AUTO& = 0
Global Const PEAC_NORMAL& = 1
Global Const PEAC_LOG& = 2

'Global Const PEMC_HIDE& = 0
'Global Const PEMC_SHOW& = 1
'Global Const PEMC_GRAYED& = 2

'Global Const PECM_SHOW& = 0
'Global Const PECM_GRAYED& = 1
'Global Const PECM_HIDE& = 2

'Global Const PECMS_UNCHECKED& = 0
'Global Const PECMS_CHECKED& = 1

'Global Const PECML_TOP& = 0
'Global Const PECML_ABOVE_SEPARATOR& = 1
'Global Const PECML_BELOW_SEPARATOR& = 2
'Global Const PECML_BOTTOM& = 3

'Global Const PEGPM_LINE& = 0
'Global Const PEGPM_BAR& = 1
'Global Const PEGPM_STICK& = 4
'Global Const PEGPM_POINT& = 2
'Global Const PEGPM_AREA& = 3
'Global Const PEGPM_AREASTACKED& = 4
'Global Const PEGPM_AREASTACKEDPERCENT& = 5
'Global Const PEGPM_BARSTACKED& = 6
'Global Const PEGPM_BARSTACKEDPERCENT& = 7
'Global Const PEGPM_POINTSPLUSBFL& = 8
'Global Const PEGPM_POINTSPLUSBFLGRAPHED& = 9
'Global Const PEGPM_HISTOGRAM& = 10
'Global Const PEGPM_SPECIFICPLOTMODE& = 11
'Global Const PEGPM_BUBBLE& = 12
'Global Const PEGPM_POINTSPLUSBFC& = 13
'Global Const PEGPM_POINTSPLUSBFCGRAPHED& = 14
'Global Const PEGPM_POINTSPLUSSPLINE& = 15
'Global Const PEGPM_SPLINE& = 16
'Global Const PEGPM_POINTSPLUSLINE& = 17
'Global Const PEGPM_HORIZONTALBAR& = 18
'Global Const PEGPM_HORZBARSTACKED& = 19
'Global Const PEGPM_HORZBARSTACKEDPERCENT& = 20
'Global Const PEGPM_STEP& = 21
'Global Const PEGPM_RIBBON& = 22
'Global Const PEGPM_CONTOURLINES& = 23
'Global Const PEGPM_CONTOURCOLORS& = 24
'Global Const PEGPM_HIGHLOWBAR& = 25
'Global Const PEGPM_HIGHLOWLINE& = 26
'Global Const PEGPM_HIGHLOWCLOSE& = 27
'Global Const PEGPM_OPENHIGHLOWCLOSE& = 28
'Global Const PEGPM_BOXPLOT& = 29

'Global Const PECPS_NONE& = 0
'Global Const PECPS_XVALUE& = 1
'Global Const PECPS_YVALUE& = 2
'Global Const PECPS_XYVALUES& = 3

'Global Const PEAUI_NONE& = 0
'Global Const PEAUI_ALL& = 1
'Global Const PEAUI_DISABLEKEYBOARD& = 2
'Global Const PEAUI_DISABLEMOUSE& = 3

Global Const PEGLC_BOTH& = 0
Global Const PEGLC_YAXIS& = 1
Global Const PEGLC_XAXIS& = 2
Global Const PEGLC_NONE& = 3

'Global Const PEAS_SUMPP& = 51
'Global Const PEAS_MINAP& = 1
'Global Const PEAS_MINPP& = 52
'Global Const PEAS_MAXAP& = 2
'Global Const PEAS_MAXPP& = 53
'Global Const PEAS_AVGAP& = 3
'Global Const PEAS_AVGPP& = 54
'Global Const PEAS_P1SDAP& = 4
'Global Const PEAS_P1SDPP& = 55
'Global Const PEAS_P2SDAP& = 5
'Global Const PEAS_P2SDPP& = 56
'Global Const PEAS_P3SDAP& = 6
'Global Const PEAS_P3SDPP& = 57
'Global Const PEAS_M1SDAP& = 7
'Global Const PEAS_M1SDPP& = 58
'Global Const PEAS_M2SDAP& = 8
'Global Const PEAS_M2SDPP& = 59
'Global Const PEAS_M3SDAP& = 9
'Global Const PEAS_M3SDPP& = 60
'Global Const PEAS_PARETO_ASC& = 90
'Global Const PEAS_PARETO_DEC& = 91

'Global Const PEPTGI_FIRSTPOINTS& = 0
'Global Const PEPTGI_LASTPOINTS& = 1

'Global Const PEPTGV_SEQUENTIAL& = 0
'Global Const PEPTGV_RANDOM& = 1

'Global Const PEGPT_GRAPH& = 0
'Global Const PEGPT_TABLE& = 1
'Global Const PEGPT_BOTH& = 2

'Global Const PETW_GRAPHED& = 0
'Global Const PETW_ALLSUBSETS& = 1

'Global Const PEDLT_PERCENTAGE& = 0
'Global Const PEDLT_VALUE& = 1

Global Const PEMSC_NONE& = 0
Global Const PEMSC_MIN& = 1
Global Const PEMSC_MAX& = 2
Global Const PEMSC_MINMAX& = 3

'Global Const IDEXPORTBUTTON& = 1015
'Global Const IDMAXIMIZEBUTTON& = 1016
'Global Const IDORIGINALBUTTON& = 1109

'Global Const PEHS_NONE& = 0
'Global Const PEHS_SUBSET& = 1
'Global Const PEHS_POINT& = 2
'Global Const PEHS_GRAPH& = 3
'Global Const PEHS_TABLE& = 4
'Global Const PEHS_DATAPOINT& = 5
'Global Const PEHS_ANNOTATION& = 6
'Global Const PEHS_XAXISANNOTATION& = 7
'Global Const PEHS_YAXISANNOTATION& = 8
'Global Const PEHS_HORZLINEANNOTATION& = 9
'Global Const PEHS_VERTLINEANNOTATION& = 10
'Global Const PEHS_MAINTITLE& = 11
'Global Const PEHS_SUBTITLE& = 12
'Global Const PEHS_MULTISUBTITLE& = 13
'Global Const PEHS_MULTIBOTTOMTITLE& = 14
'Global Const PEHS_YAXISLABEL& = 15
'Global Const PEHS_XAXISLABEL& = 16
'Global Const PEHS_YAXIS& = 17
'Global Const PEHS_XAXIS& = 18
'Global Const PEHS_YAXISGRIDNUMBER& = 19
'Global Const PEHS_RYAXISGRIDNUMBER& = 20
'Global Const PEHS_XAXISGRIDNUMBER& = 21
'Global Const PEHS_TXAXISGRIDNUMBER& = 22
'Global Const PEHS_TABLEANNOTATION& = 23
'Global Const PEHS_TABLEANNOTATION19& = 42

'Global Const PESPM_NONE& = 0
'Global Const PESPM_HIGHLOWBAR& = 1
'Global Const PESPM_HIGHLOWLINE& = 2
'Global Const PESPM_HIGHLOWCLOSE& = 3
'Global Const PESPM_OPENHIGHLOWCLOSE& = 4
'Global Const PESPM_BOXPLOT& = 5

'Global Const PEZIO_NORMAL& = 0
'Global Const PEZIO_RECT& = 1
'Global Const PEZIO_LINE& = 2

'Global Const PETS_GRIDSTYLE& = 0
'Global Const PETS_THICK& = 1
'Global Const PETS_DOT& = 2
'Global Const PETS_DASH& = 3
'Global Const PETS_1UNIT& = 4
'Global Const PETS_THIN& = 5

Global Const PEZS_FRAMED_RECT& = 0
Global Const PEZS_RO2_NOT& = 1

'Global Const PECPL_TOP_LEFT& = 0
'Global Const PECPL_TOP_RIGHT& = 1

'Global Const PELS_2_LINE& = 0
'Global Const PELS_1_LINE& = 1
'Global Const PELS_1_LINE_INSIDE_AXIS& = 2
'Global Const PELS_1_LINE_TOP_OF_AXIS& = 3

'Global Const PEMAS_GROUP_ALL_AXES& = 0
'Global Const PEMAS_SEPARATE_AXES& = 1

'Global Const PETAHO_HORZ& = 0
'Global Const PETAHO_90& = 1
'Global Const PETAHO_270& = 2

'Global Const PETAL_TOP_CENTER& = 0
'Global Const PETAL_TOP_LEFT& = 1
'Global Const PETAL_LEFT_CENTER& = 2
'Global Const PETAL_BOTTOM_LEFT& = 3
'Global Const PETAL_BOTTOM_CENTER& = 4
'Global Const PETAL_BOTTOM_RIGHT& = 5
'Global Const PETAL_RIGHT_CENTER& = 6
'Global Const PETAL_TOP_RIGHT& = 7
'Global Const PETAL_INSIDE_TOP_CENTER& = 8
'Global Const PETAL_INSIDE_TOP_LEFT& = 9
'Global Const PETAL_INSIDE_LEFT_CENTER& = 10
'Global Const PETAL_INSIDE_BOTTOM_LEFT& = 11
'Global Const PETAL_INSIDE_BOTTOM_CENTER& = 12
'Global Const PETAL_INSIDE_BOTTOM_RIGHT& = 13
'Global Const PETAL_INSIDE_RIGHT_CENTER& = 14
'Global Const PETAL_INSIDE_TOP_RIGHT& = 15
'Global Const PETAL_INSIDE_AXIS& = 100
'Global Const PETAL_INSIDE_AXIS_0& = 100
'Global Const PETAL_INSIDE_AXIS_1& = 101
'Global Const PETAL_INSIDE_AXIS_2& = 102
'Global Const PETAL_INSIDE_AXIS_3& = 103
'Global Const PETAL_INSIDE_AXIS_4& = 104
'Global Const PETAL_INSIDE_AXIS_5& = 105
'Global Const PETAL_OUTSIDE_AXIS& = 200
'Global Const PETAL_OUTSIDE_AXIS_0& = 200
'Global Const PETAL_OUTSIDE_AXIS_1& = 201
'Global Const PETAL_OUTSIDE_AXIS_2& = 202
'Global Const PETAL_OUTSIDE_AXIS_3& = 203
'Global Const PETAL_OUTSIDE_AXIS_4& = 204
'Global Const PETAL_OUTSIDE_AXIS_5& = 205
'Global Const PETAL_INSIDE_TABLE& = 300

Global Const PETAB_DROP_SHADOW& = 0
Global Const PETAB_SINGLE_LINE& = 1
Global Const PETAB_NO_BORDER& = 2
Global Const PETAB_INSET& = 3

'Global Const PETAAL_TOP_FULL_WIDTH& = 0
'Global Const PETAAL_TOP_LEFT& = 1
'Global Const PETAAL_TOP_CENTER& = 2
'Global Const PETAAL_TOP_RIGHT& = 3
'Global Const PETAAL_BOTTOM_FULL_WIDTH& = 4
'Global Const PETAAL_BOTTOM_LEFT& = 5
'Global Const PETAAL_BOTTOM_CENTER& = 6
'Global Const PETAAL_BOTTOM_RIGHT& = 7
'Global Const PETAAL_TOP_TABLE_SPACED& = 8
'Global Const PETAAL_BOTTOM_TABLE_SPACED& = 9
'Global Const PETAAL_NEW_ROW& = 100

'Global Const PETAJ_LEFT& = 0
'Global Const PETAJ_CENTER& = 1
'Global Const PETAJ_RIGHT& = 2

'Global Const PESTM_TICKS_INSIDE& = 0
'Global Const PESTM_TICKS_OUTSIDE& = 1
'Global Const PESTM_TICKS_HIDE& = 2

'Global Const PESPL_PERCENTPLUSLABEL& = 0
'Global Const PESPL_PERCENT& = 1
'Global Const PESPL_LABEL& = 2

'Global Const PESH_MONOCHROME& = 0
'Global Const PESH_BOTH& = 1


'// HORIZONTAL LINE ANNOTATIONS CAN BE WITH RESPECT TO RIGHT Y AXIS COORDINATES
'// BY ADDING 1000 TO THE FOLLOWING CONSTANTS
'Global Const PELT_THINSOLID& = 0
'Global Const PELT_DASH& = 1
'Global Const PELT_DOT& = 2
'Global Const PELT_DASHDOT& = 3
'Global Const PELT_DASHDOTDOT& = 4
'Global Const PELT_MEDIUMSOLID& = 5
'Global Const PELT_THICKSOLID& = 6
'Global Const PELAT_GRIDTICK& = 7
'Global Const PELAT_GRIDLINE& = 8
'Global Const PELT_MEDIUMTHINSOLID& = 9
'Global Const PELT_MEDIUMTHICKSOLID& = 10
'Global Const PELT_EXTRATHICKSOLID& = 11
'Global Const PELT_EXTRATHINSOLID& = 12
'Global Const PELT_EXTRAEXTRATHINSOLID& = 13
                                      
Global Const PEPS_SMALL& = 0
Global Const PEPS_MEDIUM& = 1
Global Const PEPS_LARGE& = 2
Global Const PEPS_MICRO& = 3

Global Const PEPT_PLUS& = 0
Global Const PEPT_CROSS& = 1
Global Const PEPT_DOT& = 2
Global Const PEPT_DOTSOLID& = 3
Global Const PEPT_SQUARE& = 4
Global Const PEPT_SQUARESOLID& = 5
Global Const PEPT_DIAMOND& = 6
Global Const PEPT_DIAMONDSOLID& = 7
Global Const PEPT_UPTRIANGLE& = 8
Global Const PEPT_UPTRIANGLESOLID& = 9
Global Const PEPT_DOWNTRIANGLE& = 10
Global Const PEPT_DOWNTRIANGLESOLID& = 11
Global Const PEPT_DASH& = 72
Global Const PEPT_PIXEL& = 73
Global Const PEPT_ARROW_N& = 92
Global Const PEPT_ARROW_NE& = 93
Global Const PEPT_ARROW_E& = 94
Global Const PEPT_ARROW_SE& = 95
Global Const PEPT_ARROW_S& = 96
Global Const PEPT_ARROW_SW& = 97
Global Const PEPT_ARROW_W& = 98
Global Const PEPT_ARROW_NW& = 99

'Global Const PEADL_NONE& = 0
'Global Const PEADL_DATAVALUES& = 1
'Global Const PEADL_POINTLABELS& = 2
'Global Const PEADL_DATAPOINTLABELS& = 3

Global Const PEAZ_NONE& = 0
Global Const PEAZ_HORIZONTAL& = 1
Global Const PEAZ_VERTICAL& = 2
Global Const PEAZ_HORZANDVERT& = 3

'Global Const PEBFD_2ND& = 0
'Global Const PEBFD_3RD& = 1
'Global Const PEBFD_4TH& = 2

'Global Const PEBS_SMALL& = 0
'Global Const PEBS_MEDIUM& = 1
'Global Const PEBS_LARGE& = 2

'Global Const PECG_COARSE& = 0
'Global Const PECG_MEDIUM& = 1
'Global Const PECG_FINE& = 2

'Global Const PEAE_NONE& = 0
'Global Const PEAE_ALLSUBSETS& = 1
'Global Const PEAE_INDSUBSETS& = 2

'Global Const PECM_NOCURSOR& = 0
'Global Const PECM_POINT& = 1
'Global Const PECM_DATACROSS& = 2
'Global Const PECM_DATASQUARE& = 3
'Global Const PECM_FLOATINGY& = 4
'Global Const PECM_FLOATINGXY& = 5
'Global Const PECM_FLOATINGXONLY& = 6
'Global Const PECM_FLOATINGYONLY& = 7

'// GRAPH ANNOTATIONS CAN BE WITH RESPECT TO RIGHT Y AXIS COORDINATES
'// BY ADDING 1000 TO THE FOLLOWING CONSTANTS
'Global Const PEGAT_NOSYMBOL& = 0
'Global Const PEGAT_PLUS& = 1
'Global Const PEGAT_CROSS& = 2
'Global Const PEGAT_DOT& = 3
'Global Const PEGAT_DOTSOLID& = 4
'Global Const PEGAT_SQUARE& = 5
'Global Const PEGAT_SQUARESOLID& = 6
'Global Const PEGAT_DIAMOND& = 7
'Global Const PEGAT_DIAMONDSOLID& = 8
'Global Const PEGAT_UPTRIANGLE& = 9
'Global Const PEGAT_UPTRIANGLESOLID& = 10
'Global Const PEGAT_DOWNTRIANGLE& = 11
'Global Const PEGAT_DOWNTRIANGLESOLID& = 12
'Global Const PEGAT_SMALLPLUS& = 13
'Global Const PEGAT_SMALLCROSS& = 14
'Global Const PEGAT_SMALLDOT& = 15
'Global Const PEGAT_SMALLDOTSOLID& = 16
'Global Const PEGAT_SMALLSQUARE& = 17
'Global Const PEGAT_SMALLSQUARESOLID& = 18
'Global Const PEGAT_SMALLDIAMOND& = 19
'Global Const PEGAT_SMALLDIAMONDSOLID& = 20
'Global Const PEGAT_SMALLUPTRIANGLE& = 21
'Global Const PEGAT_SMALLUPTRIANGLESOLID& = 22
'Global Const PEGAT_SMALLDOWNTRIANGLE& = 23
'Global Const PEGAT_SMALLDOWNTRIANGLESOLID& = 24
'Global Const PEGAT_LARGEPLUS& = 25
'Global Const PEGAT_LARGECROSS& = 26
'Global Const PEGAT_LARGEDOT& = 27
'Global Const PEGAT_LARGEDOTSOLID& = 28
'Global Const PEGAT_LARGESQUARE& = 29
'Global Const PEGAT_LARGESQUARESOLID& = 30
'Global Const PEGAT_LARGEDIAMOND& = 31
'Global Const PEGAT_LARGEDIAMONDSOLID& = 32
'Global Const PEGAT_LARGEUPTRIANGLE& = 33
'Global Const PEGAT_LARGEUPTRIANGLESOLID& = 34
'Global Const PEGAT_LARGEDOWNTRIANGLE& = 35
'Global Const PEGAT_LARGEDOWNTRIANGLESOLID& = 36

'Global Const PEGAT_POINTER& = 37

'Global Const PEGAT_THINSOLIDLINE& = 38
'Global Const PEGAT_DASHLINE& = 39
'Global Const PEGAT_DOTLINE& = 40
'Global Const PEGAT_DASHDOTLINE& = 41
'Global Const PEGAT_DASHDOTDOTLINE& = 42
'Global Const PEGAT_MEDIUMSOLIDLINE& = 43
'Global Const PEGAT_THICKSOLIDLINE& = 44
'Global Const PEGAT_LINECONTINUE& = 45

'Global Const PEGAT_TOPLEFT& = 46
'Global Const PEGAT_BOTTOMRIGHT& = 47

'Global Const PEGAT_RECT_THIN& = 48
'Global Const PEGAT_RECT_DASH& = 49
'Global Const PEGAT_RECT_DOT& = 50
'Global Const PEGAT_RECT_DASHDOT& = 51
'Global Const PEGAT_RECT_DASHDOTDOT& = 52
'Global Const PEGAT_RECT_MEDIUM& = 53
'Global Const PEGAT_RECT_THICK& = 54
'Global Const PEGAT_RECT_FILL& = 55

'Global Const PEGAT_ROUNDRECT_THIN& = 56
'Global Const PEGAT_ROUNDRECT_DASH& = 57
'Global Const PEGAT_ROUNDRECT_DOT& = 58
'Global Const PEGAT_ROUNDRECT_DASHDOT& = 59
'Global Const PEGAT_ROUNDRECT_DASHDOTDOT& = 60
'Global Const PEGAT_ROUNDRECT_MEDIUM& = 61
'Global Const PEGAT_ROUNDRECT_THICK& = 62
'Global Const PEGAT_ROUNDRECT_FILL& = 63

'Global Const PEGAT_ELLIPSE_THIN& = 64
'Global Const PEGAT_ELLIPSE_DASH& = 65
'Global Const PEGAT_ELLIPSE_DOT& = 66
'Global Const PEGAT_ELLIPSE_DASHDOT& = 67
'Global Const PEGAT_ELLIPSE_DASHDOTDOT& = 68
'Global Const PEGAT_ELLIPSE_MEDIUM& = 69
'Global Const PEGAT_ELLIPSE_THICK& = 70
'Global Const PEGAT_ELLIPSE_FILL& = 71

'Global Const PEGAT_DASH& = 72
'Global Const PEGAT_PIXEL& = 73

'Global Const PEGAT_STARTPOLY& = 74
'Global Const PEGAT_ADDPOLYPOINT& = 75
'Global Const PEGAT_ENDPOLYGON& = 76
'Global Const PEGAT_ENDPOLYLINE_THIN& = 77
'Global Const PEGAT_ENDPOLYLINE_MEDIUM& = 78
'Global Const PEGAT_ENDPOLYLINE_THICK& = 79
'Global Const PEGAT_ENDPOLYLINE_DASH& = 80
'Global Const PEGAT_ENDPOLYLINE_DOT& = 81
'Global Const PEGAT_ENDPOLYLINE_DASHDOT& = 82
'Global Const PEGAT_ENDPOLYLINE_DASHDOTDOT& = 83

'Global Const PEGAT_STARTTEXT& = 84
'Global Const PEGAT_ADDTEXT& = 85
'Global Const PEGAT_PARAGRAPH& = 86

'Global Const PEGAT_MEDIUMTHINSOLID& = 87
'Global Const PEGAT_MEDIUMTHICKSOLID& = 88
'Global Const PEGAT_EXTRATHICKSOLID& = 89
'Global Const PEGAT_EXTRATHINSOLID& = 90
'Global Const PEGAT_EXTRAEXTRATHINSOLID& = 91

'Global Const PEGAT_ARROW_N& = 92
'Global Const PEGAT_ARROW_NE& = 93
'Global Const PEGAT_ARROW_E& = 94
'Global Const PEGAT_ARROW_SE& = 95
'Global Const PEGAT_ARROW_S& = 96
'Global Const PEGAT_ARROW_SW& = 97
'Global Const PEGAT_ARROW_W& = 98
'Global Const PEGAT_ARROW_NW& = 99

'Global Const PEDTM_NONE& = 0
'Global Const PEDTM_VB& = 1
'Global Const PEDTM_DELPHI& = 2

'Global Const PESC_POLAR& = 0
'Global Const PESC_SMITH& = 1
'Global Const PESC_ROSE& = 2
'Global Const PESC_ADMITTANCE& = 3

'Global Const PESA_ALL& = 0
'Global Const PESA_AXISLABELS& = 1
'Global Const PESA_GRIDNUMBERS& = 2
'Global Const PESA_NONE& = 3
'Global Const PESA_LABELONLY& = 4
'Global Const PESA_EMPTY& = 5

'Global Const PEMPS_NONE& = 0
'Global Const PEMPS_SMALL& = 1
'Global Const PEMPS_MEDIUM& = 2
'Global Const PEMPS_LARGE& = 3

'Global Const PESS_NONE& = 0
'Global Const PESS_FINANCIAL& = 1

'Global Const PELL_TOP& = 0
'Global Const PELL_BOTTOM& = 1
'Global Const PELL_LEFT& = 2
'Global Const PELL_RIGHT& = 3

'Global Const PEHSS_SMALL& = 0
'Global Const PEHSS_MEDIUM& = 1
'Global Const PEHSS_LARGE& = 2

Global Const PEDS_NONE& = 0
Global Const PEDS_SHADOWS& = 1
Global Const PEDS_3D& = 2

'Global Const PEGS_THIN& = 0
'Global Const PEGS_THICK& = 1
'Global Const PEGS_DOT& = 2
'Global Const PEGS_DASH& = 3
'Global Const PEGS_ONEPIXEL& = 4

'Global Const PEFVP_AUTO& = 0
'Global Const PEFVP_VERT& = 1
'Global Const PEFVP_HORZ& = 2
'Global Const PEFVP_SLANTED& = 3

'Global Const PEMAS_NONE& = 0
'Global Const PEMAS_THIN& = 1
'Global Const PEMAS_MEDIUM& = 2
'Global Const PEMAS_THICK& = 3
'Global Const PEMAS_THICKPLUSTICK& = 4

'Global Const PERI_INCBY15& = 0
'Global Const PERI_INCBY10& = 1
'Global Const PERI_INCBY5& = 2
'Global Const PERI_INCBY2& = 3
'Global Const PERI_INCBY1& = 4
'Global Const PERI_DECBY1& = 5
'Global Const PERI_DECBY2& = 6
'Global Const PERI_DECBY5& = 7
'Global Const PERI_DECBY10& = 8
'Global Const PERI_DECBY15& = 9

'Global Const PERD_WIREFRAME& = 0
'Global Const PERD_PLOTTINGMETHOD& = 1
'Global Const PERD_FULLDETAIL& = 2

'Global Const PESBB_WHILEROTATING& = 0
'Global Const PESBB_ALWAYS& = 1
'Global Const PESBB_NEVER& = 2

'// PolyModes
'Global Const PEPM_SURFACEPOLYGONS& = 1
'Global Const PEPM_3DBAR& = 2
'Global Const PEPM_POLYGONDATA& = 3
'Global Const PEPM_SCATTER& = 4

'// Plotting Methods
'Global Const PEPLM_WIREFRAME& = 0
'Global Const PEPLM_SURFACE& = 1
'Global Const PEPLM_SURFACE_W_SHADING& = 2
'Global Const PEPLM_SURFACE_W_PIXELS& = 3
'Global Const PEPLM_SURFACE_W_CONTOUR& = 4

'// Plotting Methods for Scatter Graph
'Global Const PEPLM_POINTS& = 0
'Global Const PEPLM_LINES& = 1
'Global Const PEPLM_POINTS_AND_LINES& = 2

'Global Const PESC_NONE& = 0
'Global Const PESC_TOPLINES& = 1
'Global Const PESC_BOTTOMLINES& = 2
'Global Const PESC_TOPCOLORS& = 3
'Global Const PESC_BOTTOMCOLORS& = 4

'Global Const PESS_WHITESHADING& = 0
'Global Const PESS_COLORSHADING& = 1

'Global Const PEP_nOBJECTTYPE& = 2100

'Global Const PEP_szMAINTITLE& = 2105
'Global Const PEP_szSUBTITLE& = 2110
'Global Const PEP_nSUBSETS& = 2115
'Global Const PEP_nPOINTS& = 2120
'Global Const PEP_szaSUBSETLABELS& = 2125
'Global Const PEP_szaPOINTLABELS& = 2130
'Global Const PEP_faXDATA& = 2135
'Global Const PEP_faYDATA& = 2140

'Global Const PEP_bMONOWITHSYMBOLS& = 2145
'Global Const PEP_nDEFORIENTATION& = 2150
'Global Const PEP_bPREPAREIMAGES& = 2155

'Global Const PEP_b3DDIALOGS& = 2160

'Global Const PEP_bALLOWCUSTOMIZATION& = 2165
'Global Const PEP_bALLOWEXPORTING& = 2170
'Global Const PEP_bALLOWMAXIMIZATION& = 2175
'Global Const PEP_bALLOWPOPUP& = 2180
'Global Const PEP_nALLOWUSERINTERFACE& = 2185
'Global Const PEP_bALLOWUSERINTERFACE& = 2185

'Global Const PEP_dwaSUBSETCOLORS& = 2190
'Global Const PEP_dwaSUBSETSHADES& = 2195

'Global Const PEP_nPAGEWIDTH& = 2200
'Global Const PEP_nPAGEHEIGHT& = 2205
'Global Const PEP_rectLOGICALLOC& = 2210

'Global Const PEP_bDIRTY& = 2215
'Global Const PEP_bDIALOGSHOWN& = 2220

'Global Const PEP_bCUSTOM& = 2225

'Global Const PEP_nVIEWINGSTYLE& = 2230
'Global Const PEP_nCVIEWINGSTYLE& = 2235

'Global Const PEP_nDATASHADOWS& = 2240
'Global Const PEP_nCDATASHADOWS& = 2245
'Global Const PEP_bDATASHADOWS& = 2240
'Global Const PEP_bCDATASHADOWS& = 2245

'Global Const PEP_dwMONODESKCOLOR& = 2250
'Global Const PEP_dwMONOTEXTCOLOR& = 2255
'Global Const PEP_dwMONOSHADOWCOLOR& = 2260
'Global Const PEP_dwMONOGRAPHFORECOLOR& = 2265
'Global Const PEP_dwMONOGRAPHBACKCOLOR& = 2270
'Global Const PEP_dwMONOTABLEFORECOLOR& = 2275
'Global Const PEP_dwMONOTABLEBACKCOLOR& = 2280

'Global Const PEP_dwCMONODESKCOLOR& = 2285
'Global Const PEP_dwCMONOTEXTCOLOR& = 2290
'Global Const PEP_dwCMONOSHADOWCOLOR& = 2295
'Global Const PEP_dwCMONOGRAPHFORECOLOR& = 2300
'Global Const PEP_dwCMONOGRAPHBACKCOLOR& = 2305
'Global Const PEP_dwCMONOTABLEFORECOLOR& = 2310
'Global Const PEP_dwCMONOTABLEBACKCOLOR& = 2315

'Global Const PEP_dwDESKCOLOR& = 2320
'Global Const PEP_dwTEXTCOLOR& = 2325
'Global Const PEP_dwSHADOWCOLOR& = 2330
'Global Const PEP_dwGRAPHFORECOLOR& = 2335
'Global Const PEP_dwGRAPHBACKCOLOR& = 2340
'Global Const PEP_dwTABLEFORECOLOR& = 2345
'Global Const PEP_dwTABLEBACKCOLOR& = 2350

'Global Const PEP_dwCDESKCOLOR& = 2355
'Global Const PEP_dwCTEXTCOLOR& = 2360
'Global Const PEP_dwCSHADOWCOLOR& = 2365
'Global Const PEP_dwCGRAPHFORECOLOR& = 2370
'Global Const PEP_dwCGRAPHBACKCOLOR& = 2375
'Global Const PEP_dwCTABLEFORECOLOR& = 2380
'Global Const PEP_dwCTABLEBACKCOLOR& = 2385

'Global Const PEP_dwWDESKCOLOR& = 2390
'Global Const PEP_dwWTEXTCOLOR& = 2395
'Global Const PEP_dwWSHADOWCOLOR& = 2400
'Global Const PEP_dwWGRAPHFORECOLOR& = 2405
'Global Const PEP_dwWGRAPHBACKCOLOR& = 2410
'Global Const PEP_dwWTABLEFORECOLOR& = 2415
'Global Const PEP_dwWTABLEBACKCOLOR& = 2420

'Global Const PEP_nDATAPRECISION& = 2425
'Global Const PEP_nCDATAPRECISION& = 2430
'Global Const PEP_nMAXDATAPRECISION& = 2431

'Global Const PEP_nFONTSIZE& = 2435
'Global Const PEP_nCFONTSIZE& = 2440

'Global Const PEP_szMAINTITLEFONT& = 2445
'Global Const PEP_bMAINTITLEBOLD& = 2450
'Global Const PEP_bMAINTITLEITALIC& = 2455
'Global Const PEP_bMAINTITLEUNDERLINE& = 2460
'Global Const PEP_szCMAINTITLEFONT& = 2465
'Global Const PEP_bCMAINTITLEBOLD& = 2470
'Global Const PEP_bCMAINTITLEITALIC& = 2475
'Global Const PEP_bCMAINTITLEUNDERLINE& = 2480

'Global Const PEP_szSUBTITLEFONT& = 2485
'Global Const PEP_bSUBTITLEBOLD& = 2490
'Global Const PEP_bSUBTITLEITALIC& = 2495
'Global Const PEP_bSUBTITLEUNDERLINE& = 2500
'Global Const PEP_szCSUBTITLEFONT& = 2505
'Global Const PEP_bCSUBTITLEBOLD& = 2510
'Global Const PEP_bCSUBTITLEITALIC& = 2515
'Global Const PEP_bCSUBTITLEUNDERLINE& = 2520

'Global Const PEP_szLABELFONT& = 2525
'Global Const PEP_bLABELBOLD& = 2530
'Global Const PEP_bLABELITALIC& = 2535
'Global Const PEP_bLABELUNDERLINE& = 2540
'Global Const PEP_szCLABELFONT& = 2545
'Global Const PEP_bCLABELBOLD& = 2550
'Global Const PEP_bCLABELITALIC& = 2555
'Global Const PEP_bCLABELUNDERLINE& = 2560

'Global Const PEP_szTABLEFONT& = 2565
'Global Const PEP_szCTABLEFONT& = 2570

'Global Const PEP_bCACHEBMP& = 2574
'Global Const PEP_hMEMBITMAP& = 2575
'Global Const PEP_hMEMDC& = 2580

'Global Const PEP_bALLOWSUBSETHOTSPOTS& = 2600
'Global Const PEP_bALLOWPOINTHOTSPOTS& = 2605
'Global Const PEP_structHOTSPOTDATA& = 2610
'Global Const PEP_structKEYDOWNDATA& = 2612
'Global Const PEP_bAUTOIMAGERESET& = 2615
'Global Const PEP_bALLOWTITLESDIALOG& = 2616

'Global Const PEP_nCURSORMODE& = 2617
'Global Const PEP_nCURSORSUBSET& = 2618
'Global Const PEP_nCURSORPOINT& = 2619
'Global Const PEP_nCURSORPROMPTSTYLE& = 2620
'Global Const PEP_bCURSORPROMPTTRACKING& = 2621
'Global Const PEP_bMOUSECURSORCONTROL& = 2622
'Global Const PEP_bALLOWANNOTATIONCONTROL& = 2623

'Global Const PEP_naSUBSETSTOLEGEND& = 2624
'Global Const PEP_naLEGENDANNOTATIONTYPE& = 2625
'Global Const PEP_szaLEGENDANNOTATIONTEXT& = 2626
'Global Const PEP_dwaLEGENDANNOTATIONCOLOR& = 2627
'Global Const PEP_nVERTSCROLLPOS& = 2628
'Global Const PEP_bALLOWDEBUGOUTPUT& = 2629

'Global Const PEP_szaMULTISUBTITLES& = 2630
'Global Const PEP_szaMULTIBOTTOMTITLES& = 2631
'Global Const PEP_bFOCALRECT& = 2632
'Global Const PEP_fFONTSIZEGLOBALCNTL& = 2634
'Global Const PEP_fFONTSIZETITLECNTL& = 2635
'Global Const PEP_bSUBSETBYPOINT& = 2636

'Global Const PEP_ptLASTMOUSEMOVE& = 2637
'Global Const PEP_bALLOWOLEEXPORT& = 2638

'Global Const PEP_faZDATA& = 2900
'Global Const PEP_bINVALID& = 2905
'Global Const PEP_bOBJECTINSERVER& = 2910
'Global Const PEP_hwndPARENTALCONTROL& = 2915

'Global Const PEP_bPAINTING& = 2916

'Global Const PEP_hARROWCURSOR& = 2917
'Global Const PEP_hZOOMCURSOR& = 2918
'Global Const PEP_hHANDCURSOR& = 2919
'Global Const PEP_hNODROPCURSOR& = 2920

'Global Const PEP_bNOCUSTOMPARMS& = 2921
'Global Const PEP_bNOHELP& = 2922
'Global Const PEP_szHELPFILENAME& = 2923
'Global Const PEP_bALLOWTITLEHOTSPOTS& = 2924
'Global Const PEP_bALLOWSUBTITLEHOTSPOTS& = 2925
'Global Const PEP_bALLOWBOTTOMTITLEHOTSPOTS& = 2926
'Global Const PEP_nCHARSET& = 2927
'Global Const PEP_bALLOWJPEGOUTPUT& = 2928

'Global Const PEP_bALLOWPAGE1& = 2930
'Global Const PEP_bALLOWPAGE2& = 2931
'Global Const PEP_bALLOWSUBSETSPAGE& = 2932
'Global Const PEP_bALLOWPOINTSPAGE& = 2933
'Global Const PEP_bALLOWFONTPAGE& = 2934
'Global Const PEP_bALLOWCOLORPAGE& = 2935
'Global Const PEP_bALLOWSTYLEPAGE& = 2936
'Global Const PEP_bALLOWAXISPAGE& = 2937

'Global Const PEP_szXAXISLABEL& = 3000
'Global Const PEP_szYAXISLABEL& = 3005
'Global Const PEP_nVBOUNDARYTYPES& = 3010
'Global Const PEP_fUPPERBOUNDVALUE& = 3015
'Global Const PEP_fLOWERBOUNDVALUE& = 3020
'Global Const PEP_szUPPERBOUNDTEXT& = 3025
'Global Const PEP_szLOWERBOUNDTEXT& = 3030

'Global Const PEP_nINITIALSCALEFORYDATA& = 3035
'Global Const PEP_nSCALEFORYDATA& = 3040
'Global Const PEP_nYAXISSCALECONTROL& = 3045

'Global Const PEP_nMANUALSCALECONTROLY& = 3050
'Global Const PEP_fMANUALMINY& = 3055
'Global Const PEP_fMANUALMAXY& = 3060

'Global Const PEP_bNOSCROLLINGSUBSETCONTROL& = 3065

'Global Const PEP_nSCROLLINGSUBSETS& = 3070
'Global Const PEP_nCSCROLLINGSUBSETS& = 3075

'Global Const PEP_naRANDOMSUBSETSTOGRAPH& = 3080
'Global Const PEP_naCRANDOMSUBSETSTOGRAPH& = 3085

'Global Const PEP_nPLOTTINGMETHOD& = 3090
'Global Const PEP_nCPLOTTINGMETHOD& = 3095

'Global Const PEP_nGRIDLINECONTROL& = 3100
'Global Const PEP_nCGRIDLINECONTROL& = 3105

'Global Const PEP_bGRIDINFRONT& = 3110
'Global Const PEP_bCGRIDINFRONT& = 3115

'Global Const PEP_bTREATCOMPSASNORMAL& = 3120
'Global Const PEP_bCTREATCOMPSASNORMAL& = 3125

'Global Const PEP_nCOMPARISONSUBSETS& = 3130

'Global Const PEP_bALLOWCOORDPROMPTING& = 3200
'Global Const PEP_bALLOWGRAPHHOTSPOTS& = 3205
'Global Const PEP_bALLOWDATAHOTSPOTS& = 3210
'Global Const PEP_bMARKDATAPOINTS& = 3215
'Global Const PEP_bCMARKDATAPOINTS& = 3220

'Global Const PEP_nRYAXISCOMPARISONSUBSETS& = 3225
'Global Const PEP_nRYAXISSCALECONTROL& = 3230
'Global Const PEP_nINITIALSCALEFORRYDATA& = 3235
'Global Const PEP_nMANUALSCALECONTROLRY& = 3240
'Global Const PEP_fMANUALMINRY& = 3245
'Global Const PEP_fMANUALMAXRY& = 3250
'Global Const PEP_szRYAXISLABEL& = 3255
'Global Const PEP_nSCALEFORRYDATA& = 3256

'Global Const PEP_bALLOWPLOTCUSTOMIZATION& = 3260
'Global Const PEP_bNEGATIVEFROMXAXIS& = 3261
'Global Const PEP_bMANUALYAXISTICKNLINE& = 3262
'Global Const PEP_fMANUALYAXISTICK& = 3263
'Global Const PEP_fMANUALYAXISLINE& = 3264
'Global Const PEP_bMANUALRYAXISTICKNLINE& = 3265
'Global Const PEP_fMANUALRYAXISTICK& = 3266
'Global Const PEP_fMANUALRYAXISLINE& = 3267
'Global Const PEP_fNULLDATAVALUE& = 3268
'Global Const PEP_nPOINTSIZE& = 3269
'Global Const PEP_naSUBSETPOINTTYPES& = 3270
'Global Const PEP_naSUBSETLINETYPES& = 3271
'Global Const PEP_bALLOWBESTFITCURVE& = 3272
'Global Const PEP_nBESTFITDEGREE& = 3273
'Global Const PEP_bALLOWSPLINE& = 3274
'Global Const PEP_nCURVEGRANULARITY& = 3275

'Global Const PEP_faAPPENDYDATA& = 3276
'Global Const PEP_szaAPPENDPOINTLABELDATA& = 3277

'Global Const PEP_bALLOWLINE& = 3279
'Global Const PEP_bALLOWPOINT& = 3280
'Global Const PEP_bALLOWBESTFITLINE& = 3281
'Global Const PEP_nALLOWZOOMING& = 3282
'Global Const PEP_bZOOMMODE& = 3283
'Global Const PEP_fZOOMMINY& = 3284
'Global Const PEP_fZOOMMAXY& = 3285

'Global Const PEP_bFORCERIGHTYAXIS& = 3286
'Global Const PEP_bALLOWPOINTSPLUSLINE& = 3287
'Global Const PEP_bALLOWPOINTSPLUSSPLINE& = 3288
'Global Const PEP_nSYMBOLFREQUENCY& = 3289

'Global Const PEP_bSHOWANNOTATIONS& = 3290
'Global Const PEP_bCSHOWANNOTATIONS& = 3202
'Global Const PEP_dwANNOTATIONCOLOR& = 3203
'Global Const PEP_dwCANNOTATIONCOLOR& = 3204
'Global Const PEP_faGRAPHANNOTATIONX& = 3291
'Global Const PEP_faGRAPHANNOTATIONY& = 3292
'Global Const PEP_szaGRAPHANNOTATIONTEXT& = 3293
'Global Const PEP_nMAXAXISANNOTATIONCLUSTER& = 3296
'Global Const PEP_faXAXISANNOTATION& = 3297
'Global Const PEP_szaXAXISANNOTATIONTEXT& = 3298
'Global Const PEP_faYAXISANNOTATION& = 3299
'Global Const PEP_szaYAXISANNOTATIONTEXT& = 3201
'Global Const PEP_bANNOTATIONSINFRONT& = 3208
'Global Const PEP_nCURSORPAGEAMOUNT& = 3211
'Global Const PEP_fLINEGAPTHRESHOLD& = 3212

'Global Const PEP_faHORZLINEANNOTATION& = 3213
'Global Const PEP_szaHORZLINEANNOTATIONTEXT& = 3214
'Global Const PEP_naHORZLINEANNOTATIONTYPE& = 3216
'Global Const PEP_dwaHORZLINEANNOTATIONCOLOR& = 3217
'Global Const PEP_faVERTLINEANNOTATION& = 3218
'Global Const PEP_szaVERTLINEANNOTATIONTEXT& = 3219
'Global Const PEP_naVERTLINEANNOTATIONTYPE& = 3221
'Global Const PEP_dwaVERTLINEANNOTATIONCOLOR& = 3222
'Global Const PEP_bSHOWGRAPHANNOTATIONS& = 3223
'Global Const PEP_bSHOWXAXISANNOTATIONS& = 3224
'Global Const PEP_bSHOWYAXISANNOTATIONS& = 3226
'Global Const PEP_bSHOWHORZLINEANNOTATIONS& = 3227
'Global Const PEP_bSHOWVERTLINEANNOTATIONS& = 3228
'Global Const PEP_bALLOWGRAPHANNOTHOTSPOTS& = 3229
'Global Const PEP_bALLOWXAXISANNOTHOTSPOTS& = 3231
'Global Const PEP_bALLOWYAXISANNOTHOTSPOTS& = 3232
'Global Const PEP_bALLOWHORZLINEANNOTHOTSPOTS& = 3233
'Global Const PEP_bALLOWVERTLINEANNOTHOTSPOTS& = 3234
'Global Const PEP_dwaGRAPHANNOTATIONCOLOR& = 3236
'Global Const PEP_dwaXAXISANNOTATIONCOLOR& = 3237
'Global Const PEP_dwaYAXISANNOTATIONCOLOR& = 3238
'Global Const PEP_nGRAPHANNOTATIONTEXTSIZE& = 3242
'Global Const PEP_nAXESANNOTATIONTEXTSIZE& = 3243
'Global Const PEP_nLINEANNOTATIONTEXTSIZE& = 3244
'Global Const PEP_naGRAPHANNOTATIONTYPE& = 3246
'Global Const PEP_nZOOMINTERFACEONLY& = 3247
'Global Const PEP_fZOOMMINX& = 3248
'Global Const PEP_fZOOMMAXX& = 3249
'Global Const PEP_nDATAHOTSPOTLIMIT& = 3251
'Global Const PEP_nHOURGLASSTHRESHOLD& = 3252
'Global Const PEP_nHORZSCROLLPOS& = 3253
'Global Const PEP_bALLOWAREA& = 3254
'Global Const PEP_bVERTORIENT90DEGREES& = 3257
'Global Const PEP_dwaPOINTCOLORS& = 3258

'Global Const PEP_naMULTIAXESSUBSETS& = 3001
'Global Const PEP_naGRAPHANNOTATIONAXIS& = 3002
'Global Const PEP_naHORZLINEANNOTATIONAXIS& = 3003
'Global Const PEP_naYAXISANNOTATIONAXIS& = 3004
'Global Const PEP_nWORKINGAXIS& = 3006
'Global Const PEP_faMULTIAXESPROPORTIONS& = 3007
'Global Const PEP_naLEGENDANNOTATIONAXIS& = 3008

'Global Const PEP_bLOGSCALEEXPLABELS& = 3009
'Global Const PEP_nPLOTTINGMETHODII& = 3011
'Global Const PEP_nCPLOTTINGMETHODII& = 3012
'Global Const PEP_faXDATAII& = 3013
'Global Const PEP_faYDATAII& = 3014
'Global Const PEP_bUSINGXDATAII& = 3016
'Global Const PEP_bUSINGYDATAII& = 3017
'Global Const PEP_nDATETIMEMODE& = 3018
'Global Const PEP_fBARWIDTH& = 3019

'Global Const PEP_nSPECIFICPLOTMODE& = 3021
'Global Const PEP_bALLOWBAR& = 3022
'Global Const PEP_structGRAPHLOC& = 3023
'Global Const PEP_faAPPENDYDATAII& = 3024
'Global Const PEP_bYAXISONRIGHT& = 3026

'Global Const PEP_nSHOWYAXIS& = 3027
'Global Const PEP_nSHOWRYAXIS& = 3028
'Global Const PEP_nSHOWXAXIS& = 3029
'Global Const PEP_nGRIDSTYLE& = 3032
'Global Const PEP_bINVERTEDYAXIS& = 3033
'Global Const PEP_bINVERTEDRYAXIS& = 3034
'Global Const PEP_dwYAXISCOLOR& = 3036
'Global Const PEP_dwRYAXISCOLOR& = 3037
'Global Const PEP_dwXAXISCOLOR& = 3038
'Global Const PEP_fFONTSIZEAXISCNTL& = 3041
'Global Const PEP_fFONTSIZELEGENDCNTL& = 3042

'Global Const PEP_bYAXISLONGTICKS& = 3043
'Global Const PEP_bRYAXISLONGTICKS& = 3044
'Global Const PEP_nMULTIAXESSEPARATORS& = 3046
'Global Const PEP_nZOOMMINAXIS& = 3047
'Global Const PEP_nZOOMMAXAXIS& = 3048

'Global Const PEP_rectGRAPH& = 3049
'Global Const PEP_rectAXIS& = 3051
'Global Const PEP_szLEFTMARGIN& = 3052
'Global Const PEP_szTOPMARGIN& = 3053
'Global Const PEP_szRIGHTMARGIN& = 3054
'Global Const PEP_szBOTTOMMARGIN& = 3056

'Global Const PEP_bAUTOSCALEDATA& = 3057
'Global Const PEP_faBESTFITCOEFFS& = 3058
'Global Const PEP_naOVERLAPMULTIAXES& = 3059
'Global Const PEP_bNOHIDDENLINESINAREA& = 3061
'Global Const PEP_bSPECIFICPLOTMODECOLOR& = 3062
'Global Const PEP_nAUTOMINMAXPADDING& = 3063

'Global Const PEP_nLOGICALLIMIT& = 3064
'Global Const PEP_bNULLDATAGAPS& = 3066
'Global Const PEP_bALLOWSTEP& = 3067
'Global Const PEP_naSUBSETDEGREE& = 3068
'Global Const PEP_bSCROLLINGVERTZOOM& = 3069

'Global Const PEP_szAXISFORMATY& = 3071
'Global Const PEP_szAXISFORMATRY& = 3072

'Global Const PEP_fZOOMMINRY& = 3073
'Global Const PEP_fZOOMMAXRY& = 3074

'Global Const PEP_n3DTHRESHOLD& = 3076

'Global Const PEP_bXAXISLONGTICKS& = 3078
'Global Const PEP_bTXAXISLONGTICKS& = 3079

'// LATEST
'Global Const PEP_nHOTSPOTSIZE& = 3081
'Global Const PEP_nLEGENDLOCATION& = 3082
'Global Const PEP_bALLOWAXISLABELHOTSPOTS& = 3083
'Global Const PEP_bALLOWAXISHOTSPOTS& = 3084
'Global Const PEP_bAPPENDWITHNOUPDATE& = 3086
'Global Const PEP_bBESTFITFIX& = 3087
'Global Const PEP_dwBOXPLOTCOLOR& = 3088
'Global Const PEP_naGRAPHANNOTATIONHOTSPOT& = 3089
'Global Const PEP_bALLOWRIBBON& = 3091
'Global Const PEP_bNOGRIDLINEMULTIPLES& = 3092
'Global Const PEP_nSPECIALSCALINGY& = 3093
'Global Const PEP_nSPECIALSCALINGRY& = 3094
'Global Const PEP_nDELTAX& = 3096
'Global Const PEP_nDELTASPERDAY& = 3097
'Global Const PEP_fSTARTTIME& = 3098
'Global Const PEP_fENDTIME& = 3099
'Global Const PEP_nLOGTICKTHRESHOLD& = 3101
'Global Const PEP_nMINIMUMPOINTSIZE& = 3102


'Global Const PEP_naAUTOSTATSUBSETS& = 3300
'Global Const PEP_bNOSTACKEDDATA& = 3305
'Global Const PEP_nPOINTSTOGRAPHINIT& = 3310

'Global Const PEP_nPOINTSTOGRAPHVERSION& = 3315
'Global Const PEP_nCPOINTSTOGRAPHVERSION& = 3320

'Global Const PEP_nPOINTSTOGRAPH& = 3325
'Global Const PEP_nCPOINTSTOGRAPH& = 3330

'Global Const PEP_naRANDOMPOINTSTOGRAPH& = 3335
'Global Const PEP_naCRANDOMPOINTSTOGRAPH& = 3340

'Global Const PEP_nFORCEVERTICALPOINTS& = 3345
'Global Const PEP_nCFORCEVERTICALPOINTS& = 3350

'Global Const PEP_nGRAPHPLUSTABLE& = 3355
'Global Const PEP_nCGRAPHPLUSTABLE& = 3360

'Global Const PEP_nTABLEWHAT& = 3365
'Global Const PEP_nCTABLEWHAT& = 3370

'Global Const PEP_bALLOWTABLEHOTSPOTS& = 3400
'Global Const PEP_bALLOWHISTOGRAM& = 3401
'Global Const PEP_naALTFREQUENCIES& = 3403
'Global Const PEP_nTARGETPOINTSTOTABLE& = 3404
'Global Const PEP_nALTFREQTHRESHOLD& = 3405
'Global Const PEP_fMANUALSTACKEDMAXY& = 3406
'Global Const PEP_nMAXPOINTSTOGRAPH& = 3407
'Global Const PEP_bNORANDOMPOINTSTOGRAPH& = 3408
'Global Const PEP_szMANUALMAXPOINTLABEL& = 3409
'Global Const PEP_szMANUALMAXDATASTRING& = 3410
'Global Const PEP_bALLOWBESTFITLINEII& = 3413
'Global Const PEP_bALLOWBESTFITCURVEII& = 3414
'Global Const PEP_bAPPENDTOEND& = 3415
'Global Const PEP_bALLOWHORIZONTALBAR& = 3416

'Global Const PEP_nFIRSTPTLABELOFFSET& = 3417
'Global Const PEP_fMANUALSTACKEDMINY& = 3418
'Global Const PEP_bALLOWHORZBARSTACKED& = 3419

'Global Const PEP_bTABLECOMPARISONSUBSETS& = 3420

'// LATEST
'Global Const PEP_bFORMATTABLE& = 3421
'Global Const PEP_bALLOWTABLE& = 3422
'Global Const PEP_nAUTOXDATA& = 3423

'Global Const PEP_nINITIALSCALEFORXDATA& = 3600
'Global Const PEP_nSCALEFORXDATA& = 3605
'Global Const PEP_nXAXISSCALECONTROL& = 3610

'Global Const PEP_nMANUALSCALECONTROLX& = 3615
'Global Const PEP_fMANUALMINX& = 3620
'Global Const PEP_fMANUALMAXX& = 3625

'Global Const PEP_bGRAPHDATALABELS& = 3630
'Global Const PEP_bCGRAPHDATALABELS& = 3635

'Global Const PEP_bALLOWBUBBLE& = 3640
'Global Const PEP_nBUBBLESIZE& = 3641
'Global Const PEP_nALLOWDATALABELS& = 3642
'Global Const PEP_szaDATAPOINTLABELS& = 3643
'Global Const PEP_bMANUALXAXISTICKNLINE& = 3644
'Global Const PEP_fMANUALXAXISTICK& = 3645
'Global Const PEP_fMANUALXAXISLINE& = 3646
'Global Const PEP_bALLOWSTICK& = 3648
'Global Const PEP_bSCROLLINGHORZZOOM& = 3652
'Global Const PEP_bNORANDOMPOINTSTOEXPORT& = 3653
'Global Const PEP_bXAXISVERTNUMBERING& = 3654
'Global Const PEP_bENGSTATIONFORMAT& = 3655
'Global Const PEP_fNULLDATAVALUEX& = 3656
'Global Const PEP_bASSUMESEQDATA& = 3657
'Global Const PEP_faAPPENDXDATA& = 3658
'Global Const PEP_faAPPENDXDATAII& = 3659

'Global Const PEP_nTXAXISCOMPARISONSUBSETS& = 3661
'Global Const PEP_nTXAXISSCALECONTROL& = 3662
'Global Const PEP_nINITIALSCALEFORTXDATA& = 3663
'Global Const PEP_nMANUALSCALECONTROLTX& = 3664
'Global Const PEP_fMANUALMINTX& = 3665
'Global Const PEP_fMANUALMAXTX& = 3666
'Global Const PEP_szTXAXISLABEL& = 3667
'Global Const PEP_nSCALEFORTXDATA& = 3668
'Global Const PEP_bMANUALTXAXISTICKNLINE& = 3669
'Global Const PEP_fMANUALTXAXISTICK& = 3670
'Global Const PEP_fMANUALTXAXISLINE& = 3671
'Global Const PEP_bFORCETOPXAXIS& = 3672
'Global Const PEP_bXAXISONTOP& = 3673
'Global Const PEP_bINVERTEDXAXIS& = 3674
'Global Const PEP_bINVERTEDTXAXIS& = 3675
'Global Const PEP_nSHOWTXAXIS& = 3676
'Global Const PEP_dwTXAXISCOLOR& = 3677

'Global Const PEP_szAXISFORMATX& = 3678
'Global Const PEP_szAXISFORMATTX& = 3679

'// LATEST
'Global Const PEP_bALLOWCONTOURLINES& = 3680
'Global Const PEP_bALLOWCONTOURCOLORS& = 3681
'Global Const PEP_fMANUALMINZ& = 3682
'Global Const PEP_fMANUALMAXZ& = 3683
'Global Const PEP_nMANUALSCALECONTROLZ& = 3684
'Global Const PEP_fMANUALZAXISLINE& = 3685
'Global Const PEP_nCONTOURLINELABELDENSITY& = 3686
'Global Const PEP_bSPECIALDATETIMEMODE& = 3687

'Global Const PEP_bSMITHCHART& = 3800
'Global Const PEP_nSMITHCHART& = 3800
'Global Const PEP_bSMARTGRID& = 3801
'Global Const PEP_nSHOWPOLARGRID& = 3802
'Global Const PEP_nZERODEGREEOFFSET& = 3803

'Global Const PEP_nGROUPINGPERCENT& = 3900
'Global Const PEP_nCGROUPINGPERCENT& = 3905
'Global Const PEP_nDATALABELTYPE& = 3910
'Global Const PEP_nCDATALABELTYPE& = 3915
'Global Const PEP_nAUTOEXPLODE& = 3920

'Global Const PEP_szZAXISLABEL& = 4000
'Global Const PEP_nDEGREEOFROTATION& = 4001
'Global Const PEP_bALLOWROTATION& = 4002
'Global Const PEP_bAUTOROTATION& = 4003
'Global Const PEP_nROTATIONINCREMENT& = 4004
'Global Const PEP_nROTATIONDETAIL& = 4005
'Global Const PEP_bALLOWHORZSCROLLBAR& = 4006
'Global Const PEP_bALLOWVERTSCROLLBAR& = 4007
'Global Const PEP_nVIEWINGHEIGHT& = 4008

'Global Const PEP_bSURFACEPOLYGONBORDERS& = 4009
'Global Const PEP_bNOSURFACEPOLYGONBORDERS& = 4009

'Global Const PEP_nSHOWBOUNDINGBOX& = 4010
'Global Const PEP_nROTATIONSPEED& = 4011
'Global Const PEP_bADDSKIRTS& = 4012
'Global Const PEP_nPOLYMODE& = 4013
'Global Const PEP_structPOLYDATA& = 4014
'Global Const PEP_dwXZBACKCOLOR& = 4015
'Global Const PEP_dwYBACKCOLOR& = 4016
'Global Const PEP_dwZAXISCOLOR& = 4017
'Global Const PEP_nSHOWZAXIS& = 4018
'Global Const PEP_bMANUALZAXISTICKNLINE& = 4019
'Global Const PEP_fMANUALZAXISTICK& = 4020
'Global Const PEP_bZAXISLONGTICKS& = 4021
'Global Const PEP_fZDISTANCE& = 4022
'Global Const PEP_bINVERTEDZAXIS& = 4023

'Global Const PEP_nSHOWCONTOUR& = 4024
'Global Const PEP_bALLOWCONTOURCONTROL& = 4025
'Global Const PEP_bSHOWCONTOURLEGENDS& = 4026

'Global Const PEP_fMANUALCONTOURLINE& = 4027
'Global Const PEP_fMANUALCONTOURMIN& = 4028
'Global Const PEP_fMANUALCONTOURMAX& = 4029
'Global Const PEP_nMANUALCONTOURSCALECONTROL& = 4030

'Global Const PEP_nSHADINGSTYLE& = 4031
'Global Const PEP_bRESETGDICACHE& = 4032

'Global Const PEP_bSHOWZAXISLINEANNOTATIONS& = 4035
'Global Const PEP_faZAXISLINEANNOTATION& = 4036
'Global Const PEP_szaZAXISLINEANNOTATIONTEXT& = 4037
'Global Const PEP_naZAXISLINEANNOTATIONTYPE& = 4038
'Global Const PEP_dwaZAXISLINEANNOTATIONCOLOR& = 4039
'Global Const PEP_faGRAPHANNOTATIONZ& = 4040

'Global Const PEP_bANNOTATIONSONSURFACES& = 4041

'Global Const PEP_bALLOWWIREFRAME& = 4042
'Global Const PEP_bALLOWSURFACE& = 4043
'Global Const PEP_bALLOWSURFACESHADING& = 4044
'Global Const PEP_bALLOWSURFACECONTOUR& = 4045
'Global Const PEP_bALLOWSURFACEPIXEL& = 4046

'Global Const PEP_bUSINGZDATAII& = 4047
'Global Const PEP_faZDATAII& = 4048
'Global Const PEP_faAPPENDZDATA& = 4049
'Global Const PEP_fNULLDATAVALUEZ& = 4050

'Global Const PEP_nINITIALSCALEFORZDATA& = 4051
'Global Const PEP_nSCALEFORZDATA& = 4052
'Global Const PEP_faAPPENDZDATAII& = 4053
'Global Const PEP_bDEGREEPROMPTING& = 4054

'Type POINT3D
'    x As Single
'    y As Single
'    Z As Single
'End Type

'Type POLYGONDATA
'    Vertices(0 To 3) As POINT3D
'    NumberOfVertices As Long
'    PolyColor As Long
'    dwReserved1 As Long
'    dwReserved2 As Long
'End Type

Type Rect
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Type GLOBALPROPERTIES
    nObjectType As Long
    szMainTitle As String * 48
    szSubTitle As String * 48
    nSubsets As Long
    nPoints As Long
    bMonoWithSymbols As Long
    nDefOrientation As Long
    nPrepareImages As Long
    b3dDialogs As Long
    bDataShadows As Long
    bAllowCustomization As Long
    bAllowExporting As Long
    bAllowMaximization As Long
    bAllowPopup As Long
    nPageWidth As Long
    nPageHeight As Long
    rectLogicalLoc As Rect
    bCustom As Long
    nViewingStyle As Long
    nCViewingStyle As Long
    dwMonoDeskColor As Long
    dwMonoTextColor As Long
    dwMonoShadowColor As Long
    dwMonoGraphForeColor As Long
    dwMonoGraphBackColor As Long
    dwMonoTableForeColor As Long
    dwMonoTableBackColor As Long
    dwCMonoDeskColor As Long
    dwCMonoTextColor As Long
    dwCMonoShadowColor As Long
    dwCMonoGraphForeColor As Long
    dwCMonoGraphBackColor As Long
    dwCMonoTableForeColor As Long
    dwCMonoTableBackColor As Long
    dwDeskColor As Long
    dwTextColor As Long
    dwShadowColor As Long
    dwGraphForeColor As Long
    dwGraphBackColor As Long
    dwTableForeColor As Long
    dwTableBackColor As Long
    dwCDeskColor As Long
    dwCTextColor As Long
    dwCShadowColor As Long
    dwCGraphForeColor As Long
    dwCGraphBackColor As Long
    dwCTableForeColor As Long
    dwCTableBackColor As Long
    nDataPrecision As Long
    nCDataPrecision As Long
    nFontSize As Long
    nCFontSize As Long
    szMainTitleFont As String * 48
    bMainTitleBold As Long
    bMainTitleItalic As Long
    bMainTitleUnderline As Long
    szCMainTitleFont As String * 48
    bCMainTitleBold As Long
    bCMainTitleItalic As Long
    bCMainTitleUnderline As Long
    szSubTitleFont As String * 48
    bSubTitleBold As Long
    bSubTitleItalic As Long
    bSubTitleUnderline As Long
    szCSubTitleFont As String * 48
    bCSubTitleBold As Long
    bCSubTitleItalic As Long
    bCSubTitleUnderline As Long
    szLabelFont As String * 48
    bLabelBold As Long
    bLabelItalic As Long
    bLabelUnderline As Long
    szCLabelFont As String * 48
    bCLabelBold As Long
    bCLabelItalic As Long
    bCLabelUnderline As Long
    szTableFont As String * 48
    szCTableFont As String * 48
    bAllowSubsetHotSpots As Long
    bAllowPointHotSpots As Long
End Type

'Type POINTSTRUCT
'    x As Long
'    y As Long
'End Type

'Type PEFILEHDR
'   nMajVersion As Long   '// ProEssentials version number
'   nMinVersion As Long
'   nObjectType As Long
'   dwSize As Long
'   extra(0 To 7) As Long
'End Type

'Type SCROLLPARMS
'    nVmin As Long '// vertical scrollbar minimum
'    nVmax As Long '// vertical scrollbar maximum
'    nVpos As Long '// vertical scrollbar position
'    nHmin As Long '// horizontal scrollbar minimum
'    nHmax As Long '// horizontal scrollbar maximum
'    nHpos As Long '// horizontal scrollbar position
'End Type

'Type HOTSPOTDATA
'    HotSpotL As Long
'    HotSpotT As Long
'    HotSpotR As Long
'    HotSpotB As Long
'    nHotSpotType As Long
'    n1 As Long
'    n2 As Long
'End Type

'Type KEYDOWNDATA
'    nChar As Integer
'    nRepCnt As Integer
'    nFlags As Integer
'End Type

'Type GRAPHLOC
'    nAxis As Long
'    fXval As Double
'    fYval As Double
'End Type

'Type TM
'    nMonth As Long
'    nDay As Long
'    nYear As Long
'    nHour As Long
'    nMinute As Long
'    nSecond As Long
'    nWeekDay As Long
'   nYearDay As Long
'End Type

'Type EXTRAAXIS
'    nSize As Long
'    fMin As Double
'    fMax As Double
'    szLabel As String * 64
'    fManualLine As Double
'    fManualTick As Double
'    szFormat As String * 16
'    nShowAxis As Long
'    nShowTickMark As Long
'    bInvertedAxis As Integer
'    bLogScale As Integer
'    dwColor As Long
'End Type

'Type CUSTOMGRIDNUMBERS
'    nAxisType As Long   '// 0=Y, 1=RIGHT Y, 2=X, 3=TOP X
'    nAxisIndex As Long  '// only used for y and ry axes, index number relates to PEP_nWORKINGAXIS
'    dNumber As Double     '// number to format
'    szData As String * 48 '// With PEvget, default format string  ...  With PEvset, completed formatted string
'End Type

'////// API FUNCTIONS //////'
Declare Function PEsetglobal Lib "PEGRP32F.DLL" (ByVal hObject&, lpData As GLOBALPROPERTIES) As Long
Declare Function PEgetglobal Lib "PEGRP32F.DLL" (ByVal hObject&, lpData As GLOBALPROPERTIES) As Long

'Declare Function PEvset Lib "PEGRP32F.DLL" Alias "PEvsetA" (ByVal hObject&, ByVal nProperty&, lpvData As Any, ByVal nItems&) As Long
'Declare Function PEvget Lib "PEGRP32F.DLL" Alias "PEvgetA" (ByVal hObject&, ByVal nProperty&, lpvDest As Any) As Long
'Declare Function PEvsetcell Lib "PEGRP32F.DLL" Alias "PEvsetcellA" (ByVal hObject&, ByVal nProperty&, ByVal nCell&, lpvData As Any) As Long
'Declare Function PEvgetcell Lib "PEGRP32F.DLL" Alias "PEvgetcellA" (ByVal hObject&, ByVal nProperty&, ByVal nCell&, lpvDest As Any) As Long
'Declare Function PEszset Lib "PEGRP32F.DLL" Alias "PEszsetA" (ByVal hObject&, ByVal nProperty&, ByVal szData$) As Long
'Declare Function PEszget Lib "PEGRP32F.DLL" Alias "PEszgetA" (ByVal hObject&, ByVal nProperty&, ByVal szData$) As Long
'Declare Function PEnset Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nData&) As Long
'Declare Function PEnget Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nProperty&) As Long
'Declare Function PElset Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nProperty&, ByVal nData&) As Long
'Declare Function PElget Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nProperty&) As Long
'Declare Function PEcreate Lib "PEGRP32F.DLL" (ByVal nObjectType&, ByVal dwStyle&, lpRect As Rect, ByVal hParent&, ByVal nId&) As Long
'Declare Function PEdestroy Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PEload Lib "PEGRP32F.DLL" (ByVal hObject&, lphGlbl As Any) As Long
'Declare Function PEstore Lib "PEGRP32F.DLL" (ByVal hObject&, lphGlbl As Any, lpdwSize As Any) As Long
'Declare Function PEloadpartial Lib "PEGRP32F.DLL" (ByVal hObject&, lphGlbl As Any) As Long
'Declare Function PEstorepartial Lib "PEGRP32F.DLL" (ByVal hObject&, lphGlbl As Any, lpdwSize As Any) As Long
'Declare Function PEgetmeta Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PEresetimage Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nLength&, ByVal nHeight&) As Long
'Declare Function PElaunchcustomize Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PElaunchexport Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PElaunchmaximize Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PElaunchtextexport Lib "PEGRP32F.DLL" Alias "PElaunchtextexportA" (ByVal hObject&, ByVal bToFile&, ByVal lpszFilename$) As Long
'Declare Function PElaunchprintdialog Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal bFullPage&, lpPoint As POINTSTRUCT) As Long
'Declare Function PElaunchcolordialog Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PElaunchfontdialog Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PElaunchpopupmenu Lib "PEGRP32F.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
'Declare Function PEreinitialize Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PEreinitializecustoms Lib "PEGRP32F.DLL" (ByVal hObject&) As Long
'Declare Function PEgethelpcontext Lib "PEGRP32F.DLL" (ByVal hWnd&) As Long
'Declare Function PEcopymetatoclipboard Lib "PEGRP32F.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
'Declare Function PEcopymetatofile Lib "PEGRP32F.DLL" Alias "PEcopymetatofileA" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long
'Declare Function PEcopybitmaptoclipboard Lib "PEGRP32F.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
'Declare Function PEcopybitmaptofile Lib "PEGRP32F.DLL" Alias "PEcopybitmaptofileA" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long
'Declare Function PEcopyoletoclipboard Lib "PEGRP32F.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
'Declare Function PEprintgraph Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal hDC&, ByVal nWidth&, ByVal nHeight&, ByVal nOrient&) As Long
'Declare Function PEconvpixeltograph Lib "PEGRP32F.DLL" (ByVal hObject&, ByRef nAxis&, ByRef nX&, ByRef nY&, ByRef fX#, ByRef fY#, ByVal bRight&, ByVal bTop&, ByVal bVV&) As Long
'Declare Function PEreset Lib "PEGRP32F.DLL" (ByVal hObject&) As Long

'Declare Function PEgethotspot Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nX&, ByVal nY&) As Long
'Declare Function PEvsetEx Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal property&, ByVal nStartingCell&, ByVal nCellCount&, lpdata As Any, lpMemSetValue As Any) As Long
'Declare Function PEvgetEx Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal property&, ByVal nStartingCell&, ByVal nCellCount&, lpdata As Any) As Long
'Declare Function PEpartialresetimage Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nStartPoint&, ByVal nPointsToAdd&) As Long
'Declare Function PEsavetofile Lib "PEGRP32F.DLL" Alias "PEsavetofileA" (ByVal hObject&, ByVal lpFileName$) As Long
'Declare Function PEloadfromfile Lib "PEGRP32F.DLL" Alias "PEloadfromfileA" (ByVal hObject&, ByVal lpFileName$) As Long
'Declare Function PEcreatefromfile Lib "PEGRP32F.DLL" Alias "PEcreatefromfileA" (ByVal lpFileName$, ByVal hParent&, lpRect As Rect, ByVal nId&) As Long
'Declare Function PEcopyjpegtoclipboard Lib "PEGRP32F.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
'Declare Function PEcopyjpegtofile Lib "PEGRP32F.DLL" Alias "PEcopyjpegtofileA" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpszFilename$) As Long

'Declare Function PEresetimageEx Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nExtX, ByVal nExtY, ByVal nOrgX, ByVal nOrgY) As Long
'Declare Function PElaunchcustomizeEx Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal nPageID) As Long
'Declare Function PEcopypngtoclipboard Lib "PEGRP32F.DLL" (ByVal hObject&, lpPoint As POINTSTRUCT) As Long
'Declare Function PEcopypngtofile Lib "PEGRP32F.DLL" Alias "PEcopypngtofileA" (ByVal hObject&, lpPoint As POINTSTRUCT, ByVal lpFileName$) As Long
'Declare Function PEcreateserialdate Lib "PEGRP32F.DLL" (pfSerial As Double, dt As TM, ByVal nType) As Long
'Declare Function PEdecipherserialdate Lib "PEGRP32F.DLL" (pfSerial As Double, dt As TM, ByVal nType) As Long
'Declare Function PEserializetofile Lib "PEGRP32F.DLL" Alias "PEserializetofileA" (ByVal hObject&, ByVal lpFileName$) As Long
'Declare Function PEvsetcellEx Lib "PEGRP32F.DLL" Alias "PEvsetcellExA" (ByVal hObject&, ByVal nProperty&, ByVal nSub&, ByVal nPt&, lpvData As Any) As Long
'Declare Function PEvgetcellEx Lib "PEGRP32F.DLL" Alias "PEvgetcellExA" (ByVal hObject&, ByVal nProperty&, ByVal nSub&, ByVal nPt&, lpvDest As Any) As Long
'Declare Function PEplaymetafile Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal hDC As Long, ByVal hMF As Long) As Long
'Declare Function PEexporttext Lib "PEGRP32F.DLL" Alias "PEexporttextA" (ByVal hObject&, ByVal nPrecision&, ByVal nTableorList&, ByVal nExportWhat&, ByVal nExportStyle&, ByVal lpszFilename$) As Long

'Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect) As Long
'Declare Function GlobalLock Lib "Kernel32" (ByVal HGLOBAL&) As Long
'Declare Function GlobalUnlock Lib "Kernel32" (ByVal HGLOBAL&) As Long
'Declare Function GlobalAlloc Lib "Kernel32" (ByVal nHowTo&, ByVal dwSize As Long) As Long
'Declare Function GlobalFree Lib "Kernel32" (ByVal HGLOBAL&) As Long
'Declare Function hwrite Lib "Kernel32" Alias "_hwrite" (ByVal HGLOBAL&, lpdata As Any, ByVal dwSize As Long) As Long
'Declare Function hread Lib "Kernel32" Alias "_hread" (ByVal HGLOBAL&, lpdata As Any, ByVal dwSize As Long) As Long
'Declare Function OpenFile Lib "Kernel32" (ByVal lpszFilename$, lpOFstruct As Any, ByVal nAccess&) As Long
'Declare Function lclose Lib "Kernel32" Alias "_lclose" (ByVal hFile&) As Long
'Declare Function SetMapMode Lib "gdi32" (ByVal hDC&, ByVal Mode&) As Long
'Declare Function SetViewportExtEx Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, lpPoint As Any) As Long
'Declare Function SetViewportOrgEx Lib "gdi32" (ByVal hDC&, ByVal x&, ByVal y&, lpPoint As Any) As Long
'Declare Function PlayMetaFile Lib "gdi32" (ByVal hDC As Long, ByVal hMF As Long) As Long

'Declare Function UpdateWindow Lib "USER32.DLL" (ByVal hObject&) As Long
'Declare Function MoveWindow Lib "USER32.DLL" (ByVal hObject&, ByVal nX&, ByVal nY&, ByVal nWidth&, ByVal nHeight&, ByVal bPaint&) As Long
'Declare Function InvalidateRect Lib "USER32.DLL" (ByVal hWnd&, lpRect As Any, ByVal bRepaint&) As Long

'Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Declare Function SelectClipRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long) As Long
'Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
'Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
'Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

'Declare Function PEprintgraphEx Lib "PEGRP32F.DLL" (ByVal hObject&, ByVal hDC&, ByVal nWidth&, ByVal nHeight&, ByVal nOriginX&, ByVal nOriginY&) As Long
