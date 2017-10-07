Attribute VB_Name = "WASP_YangYu"

Option Explicit
Option Base 1


Rem **************************************************************************************************************

Rem ******************************************************************************
Rem 供用户调用，以设定使用的标准
Rem STDID=67 --> IFC67
Rem STDID=97 --> IAPWS-IF97
Private Declare PtrSafe PtrSafe Sub SETSTD_WASP Lib "WASPCN.dll" (ByVal StdID As Integer)
Rem ******************************************************************************
Rem 供用户调用，以获知使用的标准
Rem STDID=67 --> IFC67
Rem STDID=97 --> IAPWS-IF97
Private Declare PtrSafe Sub GETSTD_WASP Lib "WASPCN.dll" (ByRef StdID As Integer)
Rem ******************************************************************************

Rem 已知压力(MPa)，求对应饱和温度(℃)
Private Declare PtrSafe Sub P2T Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比焓(kJ/kg)
Private Declare PtrSafe Sub P2HL Lib "WASPCN.dll" (ByVal P As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比焓(kJ/kg)
Private Declare PtrSafe Sub P2HG Lib "WASPCN.dll" (ByVal P As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub P2SL Lib "WASPCN.dll" (ByVal P As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub P2SG Lib "WASPCN.dll" (ByVal P As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比容(m^3/kg)
Private Declare PtrSafe Sub P2VL Lib "WASPCN.dll" (ByVal P As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub P2VG Lib "WASPCN.dll" (ByVal P As Double, ByRef V As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和温度(℃)、饱和水比焓(kJ/kg)、饱和水比熵(kJ/(kg.℃))、饱和水比容(m^3/kg)
Private Declare PtrSafe Sub P2L Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和温度(℃)、饱和汽比焓(kJ/kg)、饱和汽比熵(kJ/(kg.℃))、饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub P2G Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CPL Lib "WASPCN.dll" (ByVal P As Double, ByRef CP As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和汽定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CPG Lib "WASPCN.dll" (ByVal P As Double, ByRef CP As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CVL Lib "WASPCN.dll" (ByVal P As Double, ByRef CV As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和汽定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CVG Lib "WASPCN.dll" (ByVal P As Double, ByRef CV As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水内能(kJ/kg)
Private Declare PtrSafe Sub P2EL Lib "WASPCN.dll" (ByVal P As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽内能(kJ/kg)
Private Declare PtrSafe Sub P2EG Lib "WASPCN.dll" (ByVal P As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水音速(m/s)
Private Declare PtrSafe Sub P2SSPL Lib "WASPCN.dll" (ByVal P As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽音速(m/s)
Private Declare PtrSafe Sub P2SSPG Lib "WASPCN.dll" (ByVal P As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水定熵指数
Private Declare PtrSafe Sub P2KSL Lib "WASPCN.dll" (ByVal P As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽定熵指数
Private Declare PtrSafe Sub P2KSG Lib "WASPCN.dll" (ByVal P As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水动力粘度(Pa.s)
Private Declare PtrSafe Sub P2ETAL Lib "WASPCN.dll" (ByVal P As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽动力粘度(Pa.s)
Private Declare PtrSafe Sub P2ETAG Lib "WASPCN.dll" (ByVal P As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水运动粘度(m^2/s)
Private Declare PtrSafe Sub P2UL Lib "WASPCN.dll" (ByVal P As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽运动粘度(m^2/s)
Private Declare PtrSafe Sub P2UG Lib "WASPCN.dll" (ByVal P As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水导热系数(W/(m.℃))
Private Declare PtrSafe Sub P2RAMDL Lib "WASPCN.dll" (ByVal P As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽导热系数(W/(m.℃))
Private Declare PtrSafe Sub P2RAMDG Lib "WASPCN.dll" (ByVal P As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水普朗特数
Private Declare PtrSafe Sub P2PRNL Lib "WASPCN.dll" (ByVal P As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽普朗特数
Private Declare PtrSafe Sub P2PRNG Lib "WASPCN.dll" (ByVal P As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水介电常数
Private Declare PtrSafe Sub P2EPSL Lib "WASPCN.dll" (ByVal P As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽介电常数
Private Declare PtrSafe Sub P2EPSG Lib "WASPCN.dll" (ByVal P As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水折射率
Private Declare PtrSafe Sub P2NL Lib "WASPCN.dll" (ByVal P As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽折射率
Private Declare PtrSafe Sub P2NG Lib "WASPCN.dll" (ByVal P As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)


Rem 已知压力(MPa)和温度(℃)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PT2H Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2S Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比容(m^3/kg)
Private Declare PtrSafe Sub PT2V Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求干度(100%)
Private Declare PtrSafe Sub PT2X Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub PT Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2CP Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2CV Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求内能(kJ/kg)
Private Declare PtrSafe Sub PT2E Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求音速(m/s)
Private Declare PtrSafe Sub PT2SSP Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef a As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定熵指数
Private Declare PtrSafe Sub PT2KS Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef K As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求动力粘度(Pa.s)
Private Declare PtrSafe Sub PT2ETA Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求运动粘度(m^2/s)
Private Declare PtrSafe Sub PT2U Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求热传导系数(W/(m.℃))
Private Declare PtrSafe Sub PT2RAMD Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求普朗特数
Private Declare PtrSafe Sub PT2PRN Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求介电常数
Private Declare PtrSafe Sub PT2EPS Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求折射率
Private Declare PtrSafe Sub PT2N Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和比焓(kJ/kg)，求温度(℃)
Private Declare PtrSafe Sub PH2T Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PH2S Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求比容(m^3/kg)
Private Declare PtrSafe Sub PH2V Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub PH2X Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PH Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求温度(℃)
Private Declare PtrSafe Sub PS2T Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)
Private Declare PtrSafe Sub PS2H Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求比容(m^3/kg)
Private Declare PtrSafe Sub PS2V Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub PS2X Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求温度(℃)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub Ps Lib "WASPCN.dll" Alias "PS" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)


Rem 已知压力(MPa)和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub PV2T Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PV2H Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求比容(m^3/kg)
Private Declare PtrSafe Sub PV2S Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub PV2X Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求温度(℃)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PV Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和干度(100%)，求温度(℃)
Private Declare PtrSafe Sub PX2T Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PX2H Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PX2S Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比容(m^3/kg)
Private Declare PtrSafe Sub PX2V Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub PX Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)

Rem 已知温度(℃)，求饱和压力(MPa)？
Private Declare PtrSafe Sub T2P Lib "WASPCN.dll" (ByVal T As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比焓(kJ/kg)
Private Declare PtrSafe Sub T2HL Lib "WASPCN.dll" (ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比焓(kJ/kg)
Private Declare PtrSafe Sub T2HG Lib "WASPCN.dll" (ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub T2SL Lib "WASPCN.dll" (ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub T2SG Lib "WASPCN.dll" (ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比容(m^3/kg)
Private Declare PtrSafe Sub T2VL Lib "WASPCN.dll" (ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub T2VG Lib "WASPCN.dll" (ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比焓(kJ/kg)、饱和水比熵(kJ/(kg.℃))、饱和水比容(m^3/kg)
Private Declare PtrSafe Sub T2L Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比焓(kJ/kg)、饱和汽比熵(kJ/(kg.℃))、饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub T2G Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CPL Lib "WASPCN.dll" (ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CPG Lib "WASPCN.dll" (ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CVL Lib "WASPCN.dll" (ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CVG Lib "WASPCN.dll" (ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水内能(kJ/kg)
Private Declare PtrSafe Sub T2EL Lib "WASPCN.dll" (ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽内能(kJ/kg)
Private Declare PtrSafe Sub T2EG Lib "WASPCN.dll" (ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水音速(m/s)
Private Declare PtrSafe Sub T2SSPL Lib "WASPCN.dll" (ByVal T As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽音速(m/s)
Private Declare PtrSafe Sub T2SSPG Lib "WASPCN.dll" (ByVal T As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定熵指数
Private Declare PtrSafe Sub T2KSL Lib "WASPCN.dll" (ByVal T As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定熵指数
Private Declare PtrSafe Sub T2KSG Lib "WASPCN.dll" (ByVal T As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水动力粘度(Pa.s)
Private Declare PtrSafe Sub T2ETAL Lib "WASPCN.dll" (ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽动力粘度(Pa.s)
Private Declare PtrSafe Sub T2ETAG Lib "WASPCN.dll" (ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水运动粘度(m^2/s)
Private Declare PtrSafe Sub T2UL Lib "WASPCN.dll" (ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽运动粘度(m^2/s)
Private Declare PtrSafe Sub T2UG Lib "WASPCN.dll" (ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水导热系数(W/(m.℃))
Private Declare PtrSafe Sub T2RAMDL Lib "WASPCN.dll" (ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽导热系数(W/(m.℃))
Private Declare PtrSafe Sub T2RAMDG Lib "WASPCN.dll" (ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水普朗特数
Private Declare PtrSafe Sub T2PRNL Lib "WASPCN.dll" (ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽普朗特数
Private Declare PtrSafe Sub T2PRNG Lib "WASPCN.dll" (ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水导电率
Private Declare PtrSafe Sub T2EPSL Lib "WASPCN.dll" (ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽介电常数
Private Declare PtrSafe Sub T2EPSG Lib "WASPCN.dll" (ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水折射率
Private Declare PtrSafe Sub T2NL Lib "WASPCN.dll" (ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽折射率
Private Declare PtrSafe Sub T2NG Lib "WASPCN.dll" (ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水表面张力(N/m)
Private Declare PtrSafe Sub T2SURFT Lib "WASPCN.dll" (ByVal T As Double, ByRef SurfT As Double, ByRef Range As Integer)


Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub TH2PLP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub TH2PHP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2P Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub TH2SLP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub TH2SHP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2S Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub TH2VLP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub TH2VHP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2V Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2XLP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2XHP Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2X Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(低压的一个值)
Private Declare PtrSafe Sub THLP Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(高压的一个值)
Private Declare PtrSafe Sub THHP Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub TS2PLP Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub TS2PHP Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2P Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub TS2HLP Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub TS2HHP Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2H Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub TS2VLP Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub TS2VHP Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2V Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub TS2X Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(低压的一个值)
Private Declare PtrSafe Sub TSLP Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(高压的一个值)
Private Declare PtrSafe Sub TSHP Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(缺省为低压的一个值)
Private Declare PtrSafe Sub Ts Lib "WASPCN.dll" Alias "TS" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)


Rem 已知温度(℃)和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub TV2P Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub TV2H Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub TV2S Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub TV2X Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求压力(MPa)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、干度(100%)
Private Declare PtrSafe Sub TV Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知温度(℃)和干度(100%)，求压力(MPa)
Private Declare PtrSafe Sub TX2P Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比焓(kJ/kg)
Private Declare PtrSafe Sub TX2H Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub TX2S Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比容(m^3/kg)
Private Declare PtrSafe Sub TX2V Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求压力(MPa)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub TX Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求压力(MPa)
Private Declare PtrSafe Sub HS2P Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求温度(℃)
Private Declare PtrSafe Sub HS2T Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求比容(m^3/kg)
Private Declare PtrSafe Sub HS2V Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub HS2X Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求压力(MPa)、温度(℃)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub HS Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub HV2P Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub HV2T Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub HV2S Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub HV2X Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、干度(100%)
Private Declare PtrSafe Sub HV Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub HX2PLP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub HX2PHP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2P Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub HX2TLP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub HX2THP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2T Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub HX2SLP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub HX2SHP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2S Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub HX2VLP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub HX2VHP Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2V Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub HXLP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub HXHP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(缺省是低压的一个值)
Rem Procedure HX(Var P,T:Double;Const H:Double;Var S,V:Double;Const X:Double;Var Range:Integer);StdCall;
Private Declare PtrSafe Sub HX Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)


Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub SV2P Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub SV2T Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub SV2H Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub SV2X Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、干度(100%)
Private Declare PtrSafe Sub SV Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub SX2PLP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(中压的一个值)
Private Declare PtrSafe Sub SX2PMP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub SX2PHP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2P Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub SX2TLP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(中压的一个值)
Private Declare PtrSafe Sub SX2TMP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub SX2THP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2T Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub SX2HLP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(中压的一个值)
Private Declare PtrSafe Sub SX2HMP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub SX2HHP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2H Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub SX2VLP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(中压的一个值)
Private Declare PtrSafe Sub SX2VMP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub SX2VHP Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2V Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub SXLP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(中压的一个值)
Private Declare PtrSafe Sub SXMP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub SXHP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)


Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub VX2PLP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(低高压的一个值)
Private Declare PtrSafe Sub VX2PHP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2P Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub VX2TLP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub VX2THP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2T Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub VX2HLP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub VX2HHP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2H Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub VX2SLP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub VX2SHP Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2S Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub VXLP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub VXHP Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub VX Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)


Rem ******************************************************************************
Rem 已知两个点的X、Y值，线性插值(输入顺序为先两个X,再两个Y)
Rem Private Declare PtrSafe Sub INST2DXX Lib "WASPCN.dll" (ByVal X1 As Double, ByVal X2 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal x As Double, ByRef y As Double)
Rem 已知两个点的X、Y值，线性插值(输入顺序为先一对X,Y ，再一对X,Y)
Rem Private Declare PtrSafe Sub INST2DXY Lib "WASPCN.dll" (ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal x As Double, ByRef y As Double)


Private Declare PtrSafe Sub ABOUT_WASP Lib "WASPCN.dll" ()
Private Declare PtrSafe Sub HELP_WASP Lib "WASPCN.dll" ()
Private Declare PtrSafe Sub COPYRIGHT_WASP Lib "WASPCN.dll" ()


Rem ****************************************************************************************************
Rem 已知压力(MPa)，求对应饱和温度(℃)
Private Declare PtrSafe Sub P2T67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比焓(kJ/kg)
Private Declare PtrSafe Sub P2HL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比焓(kJ/kg)
Private Declare PtrSafe Sub P2HG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub P2SL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub P2SG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比容(m^3/kg)
Private Declare PtrSafe Sub P2VL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub P2VG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef V As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和温度(℃)、饱和水比焓(kJ/kg)、饱和水比熵(kJ/(kg.℃))、饱和水比容(m^3/kg)
Private Declare PtrSafe Sub P2L67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和温度(℃)、饱和汽比焓(kJ/kg)、饱和汽比熵(kJ/(kg.℃))、饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub P2G67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CPL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef CP As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和汽定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CPG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef CP As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CVL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef CV As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和汽定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CVG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef CV As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水内能(kJ/kg)
Private Declare PtrSafe Sub P2EL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽内能(kJ/kg)
Private Declare PtrSafe Sub P2EG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水音速(m/s)
Private Declare PtrSafe Sub P2SSPL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽音速(m/s)
Private Declare PtrSafe Sub P2SSPG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水定熵指数
Private Declare PtrSafe Sub P2KSL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽定熵指数
Private Declare PtrSafe Sub P2KSG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水动力粘度(Pa.s)
Private Declare PtrSafe Sub P2ETAL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽动力粘度(Pa.s)
Private Declare PtrSafe Sub P2ETAG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水运动粘度(m^2/s)
Private Declare PtrSafe Sub P2UL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽运动粘度(m^2/s)
Private Declare PtrSafe Sub P2UG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水导热系数(W/(m.℃))
Private Declare PtrSafe Sub P2RAMDL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽导热系数(W/(m.℃))
Private Declare PtrSafe Sub P2RAMDG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水普朗特数
Private Declare PtrSafe Sub P2PRNL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽普朗特数
Private Declare PtrSafe Sub P2PRNG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水介电常数
Private Declare PtrSafe Sub P2EPSL67 Lib "WASPCN.dll" (ByVal P As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽介电常数
Private Declare PtrSafe Sub P2EPSG67 Lib "WASPCN.dll" (ByVal P As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水折射率
Private Declare PtrSafe Sub P2NL67 Lib "WASPCN.dll" (ByVal P As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽折射率
Private Declare PtrSafe Sub P2NG67 Lib "WASPCN.dll" (ByVal P As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)


Rem 已知压力(MPa)和温度(℃)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PT2H67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2S67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比容(m^3/kg)
Private Declare PtrSafe Sub PT2V67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求干度(100%)
Private Declare PtrSafe Sub PT2X67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub PT67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2CP67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2CV67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求内能(kJ/kg)
Private Declare PtrSafe Sub PT2E67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求音速(m/s)
Private Declare PtrSafe Sub PT2SSP67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef a As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定熵指数
Private Declare PtrSafe Sub PT2KS67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef K As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求动力粘度(Pa.s)
Private Declare PtrSafe Sub PT2ETA67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求运动粘度(m^2/s)
Private Declare PtrSafe Sub PT2U67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求热传导系数(W/(m.℃))
Private Declare PtrSafe Sub PT2RAMD67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求普朗特数
Private Declare PtrSafe Sub PT2PRN67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求介电常数
Private Declare PtrSafe Sub PT2EPS67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求折射率
Private Declare PtrSafe Sub PT2N67 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和比焓(kJ/kg)，求温度(℃)
Private Declare PtrSafe Sub PH2T67 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PH2S67 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求比容(m^3/kg)
Private Declare PtrSafe Sub PH2V67 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub PH2X67 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PH67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求温度(℃)
Private Declare PtrSafe Sub PS2T67 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)
Private Declare PtrSafe Sub PS2H67 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求比容(m^3/kg)
Private Declare PtrSafe Sub PS2V67 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub PS2X67 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求温度(℃)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PS67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)


Rem 已知压力(MPa)和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub PV2T67 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PV2H67 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求比容(m^3/kg)
Private Declare PtrSafe Sub PV2S67 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub PV2X67 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求温度(℃)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PV67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和干度(100%)，求温度(℃)
Private Declare PtrSafe Sub PX2T67 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PX2H67 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PX2S67 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比容(m^3/kg)
Private Declare PtrSafe Sub PX2V67 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub PX67 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)

Rem 已知温度(℃)，求饱和压力(MPa)？
Private Declare PtrSafe Sub T2P67 Lib "WASPCN.dll" (ByVal T As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比焓(kJ/kg)
Private Declare PtrSafe Sub T2HL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比焓(kJ/kg)
Private Declare PtrSafe Sub T2HG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub T2SL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub T2SG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比容(m^3/kg)
Private Declare PtrSafe Sub T2VL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub T2VG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比焓(kJ/kg)、饱和水比熵(kJ/(kg.℃))、饱和水比容(m^3/kg)
Private Declare PtrSafe Sub T2L67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比焓(kJ/kg)、饱和汽比熵(kJ/(kg.℃))、饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub T2G67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CPL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CPG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CVL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CVG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水内能(kJ/kg)
Private Declare PtrSafe Sub T2EL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽内能(kJ/kg)
Private Declare PtrSafe Sub T2EG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水音速(m/s)
Private Declare PtrSafe Sub T2SSPL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽音速(m/s)
Private Declare PtrSafe Sub T2SSPG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定熵指数
Private Declare PtrSafe Sub T2KSL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定熵指数
Private Declare PtrSafe Sub T2KSG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水动力粘度(Pa.s)
Private Declare PtrSafe Sub T2ETAL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽动力粘度(Pa.s)
Private Declare PtrSafe Sub T2ETAG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水运动粘度(m^2/s)
Private Declare PtrSafe Sub T2UL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽运动粘度(m^2/s)
Private Declare PtrSafe Sub T2UG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水导热系数(W/(m.℃))
Private Declare PtrSafe Sub T2RAMDL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽导热系数(W/(m.℃))
Private Declare PtrSafe Sub T2RAMDG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水普朗特数
Private Declare PtrSafe Sub T2PRNL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽普朗特数
Private Declare PtrSafe Sub T2PRNG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水导电率
Private Declare PtrSafe Sub T2EPSL67 Lib "WASPCN.dll" (ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽介电常数
Private Declare PtrSafe Sub T2EPSG67 Lib "WASPCN.dll" (ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水折射率
Private Declare PtrSafe Sub T2NL67 Lib "WASPCN.dll" (ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽折射率
Private Declare PtrSafe Sub T2NG67 Lib "WASPCN.dll" (ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水表面张力(N/m)
Private Declare PtrSafe Sub T2SURFT67 Lib "WASPCN.dll" (ByVal T As Double, ByRef SurfT As Double, ByRef Range As Integer)


Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub TH2PLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub TH2PHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2P67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub TH2SLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub TH2SHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2S67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub TH2VLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub TH2VHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2V67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2XLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2XHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2X67 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(低压的一个值)
Private Declare PtrSafe Sub THLP67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(高压的一个值)
Private Declare PtrSafe Sub THHP67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub TS2PLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub TS2PHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2P67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub TS2HLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub TS2HHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2H67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub TS2VLP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub TS2VHP67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2V67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub TS2X67 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(低压的一个值)
Private Declare PtrSafe Sub TSLP67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(高压的一个值)
Private Declare PtrSafe Sub TSHP67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)


Rem 已知温度(℃)和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub TV2P67 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub TV2H67 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub TV2S67 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub TV2X67 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求压力(MPa)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、干度(100%)
Private Declare PtrSafe Sub TV67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知温度(℃)和干度(100%)，求压力(MPa)
Private Declare PtrSafe Sub TX2P67 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比焓(kJ/kg)
Private Declare PtrSafe Sub TX2H67 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub TX2S67 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比容(m^3/kg)
Private Declare PtrSafe Sub TX2V67 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求压力(MPa)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub TX67 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求压力(MPa)
Private Declare PtrSafe Sub HS2P67 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求温度(℃)
Private Declare PtrSafe Sub HS2T67 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求比容(m^3/kg)
Private Declare PtrSafe Sub HS2V67 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub HS2X67 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求压力(MPa)、温度(℃)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub HS67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub HV2P67 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub HV2T67 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub HV2S67 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub HV2X67 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、干度(100%)
Private Declare PtrSafe Sub HV67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub HX2PLP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub HX2PHP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2P67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub HX2TLP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub HX2THP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2T67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub HX2SLP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub HX2SHP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2S67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub HX2VLP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub HX2VHP67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2V67 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub HXLP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub HXHP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(缺省是低压的一个值)
Rem Procedure HX67(Var P,T:Double;Const H:Double;Var S,V:Double;Const X:Double;Var Range:Integer);StdCall;
Private Declare PtrSafe Sub HX67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)


Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub SV2P67 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub SV2T67 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub SV2H67 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub SV2X67 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、干度(100%)
Private Declare PtrSafe Sub SV67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub SX2PLP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(中压的一个值)
Private Declare PtrSafe Sub SX2PMP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub SX2PHP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2P67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub SX2TLP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(中压的一个值)
Private Declare PtrSafe Sub SX2TMP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub SX2THP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2T67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub SX2HLP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(中压的一个值)
Private Declare PtrSafe Sub SX2HMP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub SX2HHP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2H67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub SX2VLP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(中压的一个值)
Private Declare PtrSafe Sub SX2VMP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub SX2VHP67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2V67 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub SXLP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(中压的一个值)
Private Declare PtrSafe Sub SXMP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub SXHP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)


Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub VX2PLP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(低高压的一个值)
Private Declare PtrSafe Sub VX2PHP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2P67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub VX2TLP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub VX2THP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2T67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub VX2HLP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub VX2HHP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2H67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub VX2SLP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub VX2SHP67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2S67 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub VXLP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub VXHP67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub VX67 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)


Rem ****************************************************************************************************

Rem 已知压力(MPa)，求对应饱和温度(℃)
Private Declare PtrSafe Sub P2T97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比焓(kJ/kg)
Private Declare PtrSafe Sub P2HL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比焓(kJ/kg)
Private Declare PtrSafe Sub P2HG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub P2SL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub P2SG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水比容(m^3/kg)
Private Declare PtrSafe Sub P2VL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub P2VG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef V As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和温度(℃)、饱和水比焓(kJ/kg)、饱和水比熵(kJ/(kg.℃))、饱和水比容(m^3/kg)
Private Declare PtrSafe Sub P2L97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和温度(℃)、饱和汽比焓(kJ/kg)、饱和汽比熵(kJ/(kg.℃))、饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub P2G97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CPL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef CP As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和汽定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CPG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef CP As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CVL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef CV As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和汽定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub P2CVG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef CV As Double, ByRef Range As Integer)
Rem  已知压力(MPa)，求对应饱和水内能(kJ/kg)
Private Declare PtrSafe Sub P2EL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽内能(kJ/kg)
Private Declare PtrSafe Sub P2EG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水音速(m/s)
Private Declare PtrSafe Sub P2SSPL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽音速(m/s)
Private Declare PtrSafe Sub P2SSPG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水定熵指数
Private Declare PtrSafe Sub P2KSL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽定熵指数
Private Declare PtrSafe Sub P2KSG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水动力粘度(Pa.s)
Private Declare PtrSafe Sub P2ETAL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽动力粘度(Pa.s)
Private Declare PtrSafe Sub P2ETAG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水运动粘度(m^2/s)
Private Declare PtrSafe Sub P2UL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽运动粘度(m^2/s)
Private Declare PtrSafe Sub P2UG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水导热系数(W/(m.℃))
Private Declare PtrSafe Sub P2RAMDL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽导热系数(W/(m.℃))
Private Declare PtrSafe Sub P2RAMDG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水普朗特数
Private Declare PtrSafe Sub P2PRNL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽普朗特数
Private Declare PtrSafe Sub P2PRNG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水导电系数
Private Declare PtrSafe Sub P2EPSL97 Lib "WASPCN.dll" (ByVal P As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽导电系数
Private Declare PtrSafe Sub P2EPSG97 Lib "WASPCN.dll" (ByVal P As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和水折射率
Private Declare PtrSafe Sub P2NL97 Lib "WASPCN.dll" (ByVal P As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知压力(MPa)，求对应饱和汽折射率
Private Declare PtrSafe Sub P2NG97 Lib "WASPCN.dll" (ByVal P As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)


Rem 已知压力(MPa)和温度(℃)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PT2H97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2S97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比容(m^3/kg)
Private Declare PtrSafe Sub PT2V97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求干度(100%)
Private Declare PtrSafe Sub PT2X97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub PT97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2CP97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub PT2CV97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求内能(kJ/kg)
Private Declare PtrSafe Sub PT2E97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求音速(m/s)
Private Declare PtrSafe Sub PT2SSP97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef a As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求定熵指数
Private Declare PtrSafe Sub PT2KS97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef K As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求动力粘度(Pa.s)
Private Declare PtrSafe Sub PT2ETA97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求运动粘度(m^2/s)
Private Declare PtrSafe Sub PT2U97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求热传导系数(W/(m.℃))
Private Declare PtrSafe Sub PT2RAMD97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求普朗特数
Private Declare PtrSafe Sub PT2PRN97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求介电常数
Private Declare PtrSafe Sub PT2EPS97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和温度(℃)，求折射率
Private Declare PtrSafe Sub PT2N97 Lib "WASPCN.dll" (ByVal P As Double, ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和比焓(kJ/kg)，求温度(℃)
Private Declare PtrSafe Sub PH2T97 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PH2S97 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求比容(m^3/kg)
Private Declare PtrSafe Sub PH2V97 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub PH2X97 Lib "WASPCN.dll" (ByVal P As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比焓(kJ/kg)，求温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PH97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求温度(℃)
Private Declare PtrSafe Sub PS2T97 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)
Private Declare PtrSafe Sub PS2H97 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求比容(m^3/kg)
Private Declare PtrSafe Sub PS2V97 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub PS2X97 Lib "WASPCN.dll" (ByVal P As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比熵(kJ/(kg.℃))，求温度(℃)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PS97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)


Rem 已知压力(MPa)和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub PV2T97 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PV2H97 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求比容(m^3/kg)
Private Declare PtrSafe Sub PV2S97 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub PV2X97 Lib "WASPCN.dll" (ByVal P As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和比容(m^3/kg)，求温度(℃)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub PV97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知压力(MPa)和干度(100%)，求温度(℃)
Private Declare PtrSafe Sub PX2T97 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比焓(kJ/kg)
Private Declare PtrSafe Sub PX2H97 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub PX2S97 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求比容(m^3/kg)
Private Declare PtrSafe Sub PX2V97 Lib "WASPCN.dll" (ByVal P As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知压力(MPa)和干度(100%)，求温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub PX97 Lib "WASPCN.dll" (ByVal P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)

Rem 已知温度(℃)，求饱和压力(MPa)？
Private Declare PtrSafe Sub T2P97 Lib "WASPCN.dll" (ByVal T As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比焓(kJ/kg)
Private Declare PtrSafe Sub T2HL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比焓(kJ/kg)
Private Declare PtrSafe Sub T2HG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub T2SL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub T2SG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比容(m^3/kg)
Private Declare PtrSafe Sub T2VL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub T2VG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水比焓(kJ/kg)、饱和水比熵(kJ/(kg.℃))、饱和水比容(m^3/kg)
Private Declare PtrSafe Sub T2L97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽比焓(kJ/kg)、饱和汽比熵(kJ/(kg.℃))、饱和汽比容(m^3/kg)
Private Declare PtrSafe Sub T2G97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CPL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定压比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CPG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef CP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CVL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定容比热(kJ/(kg.℃))
Private Declare PtrSafe Sub T2CVG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef CV As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水内能(kJ/kg)
Private Declare PtrSafe Sub T2EL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽内能(kJ/kg)
Private Declare PtrSafe Sub T2EG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef E As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水音速(m/s)
Private Declare PtrSafe Sub T2SSPL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽音速(m/s)
Private Declare PtrSafe Sub T2SSPG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef SSP As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水定熵指数
Private Declare PtrSafe Sub T2KSL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽定熵指数
Private Declare PtrSafe Sub T2KSG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef KS As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水动力粘度(Pa.s)
Private Declare PtrSafe Sub T2ETAL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽动力粘度(Pa.s)
Private Declare PtrSafe Sub T2ETAG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef ETA As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水运动粘度(m^2/s)
Private Declare PtrSafe Sub T2UL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽运动粘度(m^2/s)
Private Declare PtrSafe Sub T2UG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef U As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水导热系数(W/(m.℃))
Private Declare PtrSafe Sub T2RAMDL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽导热系数(W/(m.℃))
Private Declare PtrSafe Sub T2RAMDG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef RAMD As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水普朗特数
Private Declare PtrSafe Sub T2PRNL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽普朗特数
Private Declare PtrSafe Sub T2PRNG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef PRN As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水介电常数
Private Declare PtrSafe Sub T2EPSL97 Lib "WASPCN.dll" (ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽介电常数
Private Declare PtrSafe Sub T2EPSG97 Lib "WASPCN.dll" (ByVal T As Double, ByRef eps As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水折射率
Private Declare PtrSafe Sub T2NL97 Lib "WASPCN.dll" (ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和汽折射率
Private Declare PtrSafe Sub T2NG97 Lib "WASPCN.dll" (ByVal T As Double, ByVal Lamd As Double, ByRef n As Double, ByRef Range As Integer)
Rem 已知温度(℃)，求饱和水表面张力(N/m)
Private Declare PtrSafe Sub T2SURFT97 Lib "WASPCN.dll" (ByVal T As Double, ByRef SurfT As Double, ByRef Range As Integer)


Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub TH2PLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub TH2PHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2P97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub TH2SLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub TH2SHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比熵(kJ/(kg.℃))(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2S97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub TH2VLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub TH2VHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求比容(m^3/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH2V97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2XLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2XHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求干度(100%)
Private Declare PtrSafe Sub TH2X97 Lib "WASPCN.dll" (ByVal T As Double, ByVal H As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(低压的一个值)
Private Declare PtrSafe Sub THLP97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(高压的一个值)
Private Declare PtrSafe Sub THHP97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比焓(kJ/kg)，求压力(MPa)、比熵(kJ/(kg.℃))、比容(m^3/kg)、干度(100%)(缺省为低压的一个值)
Private Declare PtrSafe Sub TH97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub TS2PLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub TS2PHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2P97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub TS2HLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub TS2HHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比焓(kJ/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2H97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub TS2VLP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub TS2VHP97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求比容(m^3/kg)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS2V97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub TS2X97 Lib "WASPCN.dll" (ByVal T As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(低压的一个值)
Private Declare PtrSafe Sub TSLP97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(高压的一个值)
Private Declare PtrSafe Sub TSHP97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比熵(kJ/(kg.℃))，求压力(MPa)、比焓(kJ/kg)、比容(m^3/kg)、干度(100%)(缺省为低压的一个值)
Private Declare PtrSafe Sub TS97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)


Rem 已知温度(℃)和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub TV2P97 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub TV2H97 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub TV2S97 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub TV2X97 Lib "WASPCN.dll" (ByVal T As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知温度(℃)和比容(m^3/kg)，求压力(MPa)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、干度(100%)
Private Declare PtrSafe Sub TV97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知温度(℃)和干度(100%)，求压力(MPa)
Private Declare PtrSafe Sub TX2P97 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比焓(kJ/kg)
Private Declare PtrSafe Sub TX2H97 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub TX2S97 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求比容(m^3/kg)
Private Declare PtrSafe Sub TX2V97 Lib "WASPCN.dll" (ByVal T As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知温度(℃)和干度(100%)，求压力(MPa)、比焓(kJ/kg)、比熵(kJ/(kg.℃))、比容(m^3/kg)
Private Declare PtrSafe Sub TX97 Lib "WASPCN.dll" (ByRef P As Double, ByVal T As Double, ByRef H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求压力(MPa)
Private Declare PtrSafe Sub HS2P97 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求温度(℃)
Private Declare PtrSafe Sub HS2T97 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求比容(m^3/kg)
Private Declare PtrSafe Sub HS2V97 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求干度(100%)
Private Declare PtrSafe Sub HS2X97 Lib "WASPCN.dll" (ByVal H As Double, ByVal S As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比熵(kJ/(kg.℃))，求压力(MPa)、温度(℃)、比容(m^3/kg)、干度(100%)
Private Declare PtrSafe Sub HS97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByVal S As Double, ByRef V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub HV2P97 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub HV2T97 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求比熵(kJ/(kg.℃))
Private Declare PtrSafe Sub HV2S97 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub HV2X97 Lib "WASPCN.dll" (ByVal H As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和比容(m^3/kg)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、干度(100%)
Private Declare PtrSafe Sub HV97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub HX2PLP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub HX2PHP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2P97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub HX2TLP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub HX2THP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2T97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub HX2SLP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub HX2SHP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2S97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub HX2VLP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub HX2VHP97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub HX2V97 Lib "WASPCN.dll" (ByVal H As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub HXLP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub HXHP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比焓(kJ/kg)和干度(100%)，求压力(MPa)、温度(℃)、比熵(kJ/(kg.℃))、比容(m^3/kg)(缺省是低压的一个值)
Rem Procedure HX97(Var P,T:Double;Const H:Double;Var S,V:Double;Const X:Double;Var Range:Integer);StdCall;
Private Declare PtrSafe Sub HX97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByVal H As Double, ByRef S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)


Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求压力(MPa)
Private Declare PtrSafe Sub SV2P97 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求温度(℃)
Private Declare PtrSafe Sub SV2T97 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求比焓(kJ/kg)
Private Declare PtrSafe Sub SV2H97 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求干度(100%)
Private Declare PtrSafe Sub SV2X97 Lib "WASPCN.dll" (ByVal S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和比容(m^3/kg)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、干度(100%)
Private Declare PtrSafe Sub SV97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByVal V As Double, ByRef X As Double, ByRef Range As Integer)

Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub SX2PLP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(中压的一个值)
Private Declare PtrSafe Sub SX2PMP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(高压的一个值)
Private Declare PtrSafe Sub SX2PHP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2P97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub SX2TLP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(中压的一个值)
Private Declare PtrSafe Sub SX2TMP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub SX2THP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2T97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub SX2HLP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(中压的一个值)
Private Declare PtrSafe Sub SX2HMP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub SX2HHP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比焓(kJ/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2H97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub SX2VLP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(中压的一个值)
Private Declare PtrSafe Sub SX2VMP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub SX2VHP97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX2V97 Lib "WASPCN.dll" (ByVal S As Double, ByVal X As Double, ByRef V As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(低压的一个值)
Private Declare PtrSafe Sub SXLP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(中压的一个值)
Private Declare PtrSafe Sub SXMP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(高压的一个值)
Private Declare PtrSafe Sub SXHP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比熵(kJ/(kg.℃))和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比容(m^3/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub SX97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByVal S As Double, ByRef V As Double, ByVal X As Double, ByRef Range As Integer)


Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(低压的一个值)
Private Declare PtrSafe Sub VX2PLP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(低高压的一个值)
Private Declare PtrSafe Sub VX2PHP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2P97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef P As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(低压的一个值)
Private Declare PtrSafe Sub VX2TLP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(高压的一个值)
Private Declare PtrSafe Sub VX2THP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求温度(℃)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2T97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef T As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(低压的一个值)
Private Declare PtrSafe Sub VX2HLP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(高压的一个值)
Private Declare PtrSafe Sub VX2HHP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比焓(kJ/kg)(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2H97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef H As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub VX2SLP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub VX2SHP97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub VX2S97 Lib "WASPCN.dll" (ByVal V As Double, ByVal X As Double, ByRef S As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(低压的一个值)
Private Declare PtrSafe Sub VXLP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(高压的一个值)
Private Declare PtrSafe Sub VXHP97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)
Rem 已知比容(m^3/kg)和干度(100%)，求压力(MPa)、温度(℃)、比焓(kJ/kg)、比熵(kJ/(kg.℃))(缺省是低压的一个值)
Private Declare PtrSafe Sub VX97 Lib "WASPCN.dll" (ByRef P As Double, ByRef T As Double, ByRef H As Double, ByRef S As Double, ByVal V As Double, ByVal X As Double, ByRef Range As Integer)

Rem *********************************************************************************************************************

Rem 获取当前正在使用的计算标准为（IFC67或IAPWS-IF97）
Function WASP_GetStd(ByVal IDlong As Integer) As String
    Dim CurStdID As Integer
    Call GETSTD_WASP(CurStdID)
    If IDlong = 1 Then
        If CurStdID = 67 Then
            WASP_GetStd = "IFC67"
        Else
            WASP_GetStd = "IAPWS-IF97"
        End If
    Else
        If CurStdID = 67 Then
            WASP_GetStd = "67"
        Else
            WASP_GetStd = "97"
        End If
    End If
End Function

Rem 将当前计算标准设定为IFC67或IAPWS-IF97
Function WASP_SetStd(ByVal StdID As Integer) As String
    Dim CurStdID As Integer
    Call SETSTD_WASP(StdID)
    Rem 返回当前设定的标准
    WASP_SetStd = WASP_GetStd(1)
End Function


Function T_P(ByVal P As Double)
Rem Attribute T_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和温度T(℃)?"
Rem Attribute T_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2T(P, T, Range)
    If Range = 0 Then
        T_P = "Error!"
    Else
        T_P = T
    End If
End Function


Function HL_P(ByVal P As Double)
Rem Attribute Hw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比焓Hw(kJ/kg)?"
Rem Attribute Hw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2HL(P, H, Range)
    If Range = 0 Then
        HL_P = "Error!"
    Else
        HL_P = H
    End If
End Function

Function HG_P(ByVal P As Double)
Rem Attribute Hs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比焓Hs(kJ/kg)?"
Rem Attribute Hs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2HG(P, H, Range)
    If Range = 0 Then
        HG_P = "Error!"
    Else
        HG_P = H
    End If
End Function

Function SL_P(ByVal P As Double)
Rem Attribute Sw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比熵Sw( kJ/(kg.℃) )?"
Rem Attribute Sw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SL(P, S, Range)
    If Range = 0 Then
        SL_P = "Error!"
    Else
        SL_P = S
    End If
End Function

Function SG_P(ByVal P As Double)
Rem Attribute Ss_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比熵Ss( kJ/(kg.℃) )?"
Rem Attribute Ss_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SG(P, S, Range)
    If Range = 0 Then
        SG_P = "Error!"
    Else
        SG_P = S
    End If
End Function


Function VL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2VL(P, V, Range)
    If Range = 0 Then
        VL_P = "Error!"
    Else
        VL_P = V
    End If
End Function

Function VG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2VG(P, V, Range)
    If Range = 0 Then
        VG_P = "Error!"
    Else
        VG_P = V
    End If
End Function


Function CPL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CPL(P, CP, Range)
    If Range = 0 Then
        CPL_P = "Error!"
    Else
        CPL_P = CP
    End If
End Function

Function CPG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CPG(P, CP, Range)
    If Range = 0 Then
        CPG_P = "Error!"
    Else
        CPG_P = CP
    End If
End Function

Function CVL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CVL(P, CV, Range)
    If Range = 0 Then
        CVL_P = "Error!"
    Else
        CVL_P = CV
    End If
End Function

Function CVG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CVG(P, CV, Range)
    If Range = 0 Then
        CVG_P = "Error!"
    Else
        CVG_P = CV
    End If
End Function

Function EL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EL(P, E, Range)
    If Range = 0 Then
        EL_P = "Error!"
    Else
        EL_P = E
    End If
End Function

Function EG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EG(P, E, Range)
    If Range = 0 Then
        EG_P = "Error!"
    Else
        EG_P = E
    End If
End Function

Function SSPL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SSPL(P, SSP, Range)
    If Range = 0 Then
        SSPL_P = "Error!"
    Else
        SSPL_P = SSP
    End If
End Function

Function SSPG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SSPG(P, SSP, Range)
    If Range = 0 Then
        SSPG_P = "Error!"
    Else
        SSPG_P = SSP
    End If
End Function

Function KSL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2KSL(P, KS, Range)
    If Range = 0 Then
        KSL_P = "Error!"
    Else
        KSL_P = KS
    End If
End Function

Function KSG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2KSG(P, KS, Range)
    If Range = 0 Then
        KSG_P = "Error!"
    Else
        KSG_P = KS
    End If
End Function


Function ETAL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2ETAL(P, ETA, Range)
    If Range = 0 Then
        ETAL_P = "Error!"
    Else
        ETAL_P = ETA
    End If
End Function

Function ETAG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2ETAG(P, ETA, Range)
    If Range = 0 Then
        ETAG_P = "Error!"
    Else
        ETAG_P = ETA
    End If
End Function

Function UL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2UL(P, U, Range)
    If Range = 0 Then
        UL_P = "Error!"
    Else
        UL_P = U
    End If
End Function

Function UG_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2UG(P, U, Range)
    If Range = 0 Then
        UG_P = "Error!"
    Else
        UG_P = U
    End If
End Function

Function RAMDL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2RAMDL(P, RAMD, Range)
    If Range = 0 Then
        RAMDL_P = "Error!"
    Else
        RAMDL_P = RAMD
    End If
End Function

Function RAMDG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2RAMDG(P, RAMD, Range)
    If Range = 0 Then
        RAMDG_P = "Error!"
    Else
        RAMDG_P = RAMD
    End If
End Function


Function PRNL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2PRNL(P, PRN, Range)
    If Range = 0 Then
        PRNL_P = "Error!"
    Else
        PRNL_P = PRN
    End If
End Function

Function PRNG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2PRNG(P, PRN, Range)
    If Range = 0 Then
        PRNG_P = "Error!"
    Else
        PRNG_P = PRN
    End If
End Function


Function EPSL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EPSL(P, eps, Range)
    If Range = 0 Then
        EPSL_P = "Error!"
    Else
        EPSL_P = eps
    End If
End Function

Function EPSG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EPSG(P, eps, Range)
    If Range = 0 Then
        EPSG_P = "Error!"
    Else
        EPSG_P = eps
    End If
End Function

Function NL_P(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2NL(P, Lamd, n, Range)
    If Range = 0 Then
        NL_P = "Error!"
    Else
        NL_P = n
    End If
End Function

Function NG_P(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2NG(P, Lamd, n, Range)
    If Range = 0 Then
        NG_P = "Error!"
    Else
        NG_P = n
    End If
End Function

Function H_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute H_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2H(P, T, H, Range)
    If Range = 0 Then
        H_PT = "Error!"
    Else
        H_PT = H
    End If
End Function
Function S_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute S_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2S(P, T, S, Range)
    If Range = 0 Then
        S_PT = "Error!"
    Else
        S_PT = S
    End If
End Function
Function V_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute V_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2V(P, T, V, Range)
    If Range = 0 Then
        V_PT = "Error!"
    Else
        V_PT = V
    End If
End Function
Function X_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute X_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的干度X(100%)?"
Rem Attribute X_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2X(P, T, X, Range)
    If Range = 0 Then
        X_PT = "Error!"
    Else
        X_PT = X
    End If
End Function
Function ETA_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2ETA(P, T, ETA, Range)
    If Range = 0 Then
        ETA_PT = "Error!"
    Else
        ETA_PT = ETA
    End If
End Function

Function U_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2U(P, T, U, Range)
    If Range = 0 Then
        U_PT = "Error!"
    Else
        U_PT = U
    End If
End Function


Function RAMD_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute RAMD_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的导热系数Ramd( mW/(m.℃) )?"
Rem Attribute RAMD_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求RAMD(导热系数)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2RAMD(P, T, RAMD, Range)
    If Range = 0 Then
        RAMD_PT = "Error!"
    Else
        RAMD_PT = RAMD
    End If
End Function

Function CP_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2CP(P, T, CP, Range)
    If Range = 0 Then
        CP_PT = "Error!"
    Else
        CP_PT = CP
    End If
End Function

Function CV_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2CV(P, T, CV, Range)
    If Range = 0 Then
        CV_PT = "Error!"
    Else
        CV_PT = CV
    End If
End Function

Function E_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2E(P, T, E, Range)
    If Range = 0 Then
        E_PT = "Error!"
    Else
        E_PT = E
    End If
End Function
Function KS_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute K_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的绝热指数K(100%)?"
Rem Attribute K_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求K(绝热指数)。
    Dim K As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2KS(P, T, K, Range)
    If Range = 0 Then
        KS_PT = "Error!"
    Else
        KS_PT = K
    End If
End Function

Function SSP_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute A_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的音速A (m/s)?"
Rem Attribute A_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求A(音速)。
    Dim a As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2SSP(P, T, a, Range)
    If Range = 0 Then
        SSP_PT = "Error!"
    Else
        SSP_PT = a
    End If
End Function

Function PRN_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2PRN(P, T, PRN, Range)
    If Range = 0 Then
        PRN_PT = "Error!"
    Else
        PRN_PT = PRN
    End If
End Function

Function EPS_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2EPS(P, T, eps, Range)
    If Range = 0 Then
        EPS_PT = "Error!"
    Else
        EPS_PT = eps
    End If
End Function

Function N_PT(ByVal P As Double, ByVal T As Double, ByVal Lamd As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2N(P, T, Lamd, n, Range)
    If Range = 0 Then
        N_PT = "Error!"
    Else
        N_PT = n
    End If
End Function

Function T_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute T_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2T(P, H, T, Range)
    If Range = 0 Then
        T_PH = "Error!"
    Else
        T_PH = T
    End If
End Function
Function S_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute S_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2S(P, H, S, Range)
    If Range = 0 Then
        S_PH = "Error!"
    Else
        S_PH = S
    End If
End Function
Function V_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute V_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2V(P, H, V, Range)
    If Range = 0 Then
        V_PH = "Error!"
    Else
        V_PH = V
    End If
End Function
Function X_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute X_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2X(P, H, X, Range)
    If Range = 0 Then
        X_PH = "Error!"
    Else
        X_PH = X
    End If
End Function


Function T_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute T_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的温度T(℃)?"
Rem Attribute T_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2T(P, S, T, Range)
    If Range = 0 Then
        T_PS = "Error!"
    Else
        T_PS = T
    End If
End Function
Function H_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute H_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2H(P, S, H, Range)
    If Range = 0 Then
        H_PS = "Error!"
    Else
        H_PS = H
    End If
End Function
Function V_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute V_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2V(P, S, V, Range)
    If Range = 0 Then
        V_PS = "Error!"
    Else
        V_PS = V
    End If
End Function
Function X_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute X_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2X(P, S, X, Range)
    If Range = 0 Then
        X_PS = "Error!"
    Else
        X_PS = X
    End If
End Function


Function T_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute T_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2T(P, V, T, Range)
    If Range = 0 Then
        T_PV = "Error!"
    Else
        T_PV = T
    End If
End Function
Function H_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute H_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2H(P, V, H, Range)
    If Range = 0 Then
        H_PV = "Error!"
    Else
        H_PV = H
    End If
End Function
Function S_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute S_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2S(P, V, S, Range)
    If Range = 0 Then
        S_PV = "Error!"
    Else
        S_PV = S
    End If
End Function
Function X_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute X_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),求对应的干度X(100%)?"
Rem Attribute X_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2X(P, V, X, Range)
    If Range = 0 Then
        X_PV = "Error!"
    Else
        X_PV = X
    End If
End Function
Function T_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute T_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2T(P, X, T, Range)
    If Range = 0 Then
        T_PX = "Error!"
    Else
        T_PX = T
    End If
End Function
Function H_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute H_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2H(P, X, H, Range)
    If Range = 0 Then
        H_PX = "Error!"
    Else
        H_PX = H
    End If
End Function
Function S_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute S_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2S(P, X, S, Range)
    If Range = 0 Then
        S_PX = "Error!"
    Else
        S_PX = S
    End If
End Function
Function V_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute V_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2V(P, X, V, Range)
    If Range = 0 Then
        V_PX = "Error!"
    Else
        V_PX = V
    End If
End Function


Function P_T(ByVal T As Double)
Rem Attribute P_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和压力P(MPa)?"
Rem Attribute P_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2P(T, P, Range)
    If Range = 0 Then
        P_T = "Error!"
    Else
        P_T = P
    End If
End Function
Function HL_T(ByVal T As Double)
Rem Attribute Hw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比焓Hw(kJ/kg)?"
Rem Attribute Hw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2HL(T, H, Range)
    If Range = 0 Then
        HL_T = "Error!"
    Else
        HL_T = H
    End If
End Function
Function HG_T(ByVal T As Double)
Rem Attribute Hs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比焓Hs(kJ/kg)?"
Rem Attribute Hs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HS(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2HG(T, H, Range)
    If Range = 0 Then
        HG_T = "Error!"
    Else
        HG_T = H
    End If
End Function
Function SG_T(ByVal T As Double)
Rem Attribute Ss_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比熵Ss( kJ/(kg.℃) )?"
Rem Attribute Ss_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SS(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SG(T, S, Range)
    If Range = 0 Then
        SG_T = "Error!"
    Else
        SG_T = S
    End If
End Function
Function SL_T(ByVal T As Double)
Rem Attribute Sw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比熵Sw( kJ/(kg.℃) )?"
Rem Attribute Sw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SL(T, S, Range)
    If Range = 0 Then
        SL_T = "Error!"
    Else
        SL_T = S
    End If
End Function
Function VL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2VL(T, V, Range)
    If Range = 0 Then
        VL_T = "Error!"
    Else
        VL_T = V
    End If
End Function
Function VG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2VG(T, V, Range)
    If Range = 0 Then
        VG_T = "Error!"
    Else
        VG_T = V
    End If
End Function


Function CPL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CPL(T, CP, Range)
    If Range = 0 Then
        CPL_T = "Error!"
    Else
        CPL_T = CP
    End If
End Function
Function CPG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CPG(T, CP, Range)
    If Range = 0 Then
        CPG_T = "Error!"
    Else
        CPG_T = CP
    End If
End Function


Function CVL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CVL(T, CV, Range)
    If Range = 0 Then
        CVL_T = "Error!"
    Else
        CVL_T = CV
    End If
End Function
Function CVG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CVG(T, CV, Range)
    If Range = 0 Then
        CVG_T = "Error!"
    Else
        CVG_T = CV
    End If
End Function

Function EL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EL(T, E, Range)
    If Range = 0 Then
        EL_T = "Error!"
    Else
        EL_T = E
    End If
End Function
Function EG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EG(T, E, Range)
    If Range = 0 Then
        EG_T = "Error!"
    Else
        EG_T = E
    End If
End Function

Function SSPL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SSPL(T, SSP, Range)
    If Range = 0 Then
        SSPL_T = "Error!"
    Else
        SSPL_T = SSP
    End If
End Function
Function SSPG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SSPG(T, SSP, Range)
    If Range = 0 Then
        SSPG_T = "Error!"
    Else
        SSPG_T = SSP
    End If
End Function



Function KSL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2KSL(T, KS, Range)
    If Range = 0 Then
        KSL_T = "Error!"
    Else
        KSL_T = KS
    End If
End Function
Function KSG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2KSG(T, KS, Range)
    If Range = 0 Then
        KSG_T = "Error!"
    Else
        KSG_T = KS
    End If
End Function


Function ETAL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2ETAL(T, ETA, Range)
    If Range = 0 Then
        ETAL_T = "Error!"
    Else
        ETAL_T = ETA
    End If
End Function
Function ETAG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2ETAG(T, ETA, Range)
    If Range = 0 Then
        ETAG_T = "Error!"
    Else
        ETAG_T = ETA
    End If
End Function

Function UL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2UL(T, U, Range)
    If Range = 0 Then
        UL_T = "Error!"
    Else
        UL_T = U
    End If
End Function

Function UG_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2UG(T, U, Range)
    If Range = 0 Then
        UG_T = "Error!"
    Else
        UG_T = U
    End If
End Function

Function RAMDL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2RAMDL(T, RAMD, Range)
    If Range = 0 Then
        RAMDL_T = "Error!"
    Else
        RAMDL_T = RAMD
    End If
End Function
Function RAMDG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2RAMDG(T, RAMD, Range)
    If Range = 0 Then
        RAMDG_T = "Error!"
    Else
        RAMDG_T = RAMD
    End If
End Function




Function PRNL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2PRNL(T, PRN, Range)
    If Range = 0 Then
        PRNL_T = "Error!"
    Else
        PRNL_T = PRN
    End If
End Function
Function PRNG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2PRNG(T, PRN, Range)
    If Range = 0 Then
        PRNG_T = "Error!"
    Else
        PRNG_T = PRN
    End If
End Function

Function EPSL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EPSL(T, eps, Range)
    If Range = 0 Then
        EPSL_T = "Error!"
    Else
        EPSL_T = eps
    End If
End Function
Function EPSG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EPSG(T, eps, Range)
    If Range = 0 Then
        EPSG_T = "Error!"
    Else
        EPSG_T = eps
    End If
End Function

Function NL_T(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2NL(T, Lamd, n, Range)
    If Range = 0 Then
        NL_T = "Error!"
    Else
        NL_T = n
    End If
End Function

Function NG_T(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2NG(T, Lamd, n, Range)
    If Range = 0 Then
        NG_T = "Error!"
    Else
        NG_T = n
    End If
End Function

Function SurfT_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SurfT As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SURFT(T, SurfT, Range)
    If Range = 0 Then
        SurfT_T = "Error!"
    Else
        SurfT_T = SurfT
    End If
End Function

Function P_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2P(T, H, P, Range)
    If Range = 0 Then
        P_TH = "Error!"
    Else
        P_TH = P
    End If
End Function

Function PLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2PLP(T, H, P, Range)
    If Range = 0 Then
        PLP_TH = "Error!"
    Else
        PLP_TH = P
    End If
End Function



Function PHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2PHP(T, H, P, Range)
    If Range = 0 Then
        PHP_TH = "Error!"
    Else
        PHP_TH = P
    End If
End Function

Function S_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2S(T, H, S, Range)
    If Range = 0 Then
        S_TH = "Error!"
    Else
        S_TH = S
    End If
End Function

Function SLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2SLP(T, H, S, Range)
    If Range = 0 Then
        SLP_TH = "Error!"
    Else
        SLP_TH = S
    End If
End Function

Function SHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2SHP(T, H, S, Range)
    If Range = 0 Then
        SHP_TH = "Error!"
    Else
        SHP_TH = S
    End If
End Function


Function V_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2V(T, H, V, Range)
    If Range = 0 Then
        V_TH = "Error!"
    Else
        V_TH = V
    End If
End Function


Function VLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2VLP(T, H, V, Range)
    If Range = 0 Then
        VLP_TH = "Error!"
    Else
        VLP_TH = V
    End If
End Function


Function VHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2VHP(T, H, V, Range)
    If Range = 0 Then
        VHP_TH = "Error!"
    Else
        VHP_TH = V
    End If
End Function

Function XLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2XLP(T, H, X, Range)
    If Range = 0 Then
        XLP_TH = "Error!"
    Else
        XLP_TH = X
    End If
End Function
Function XHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2XHP(T, H, X, Range)
    If Range = 0 Then
        XHP_TH = "Error!"
    Else
        XHP_TH = X
    End If
End Function
Function X_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2X(T, H, X, Range)
    If Range = 0 Then
        X_TH = "Error!"
    Else
        X_TH = X
    End If
End Function


Function PLP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2PLP(T, S, P, Range)
    If Range = 0 Then
        PLP_TS = "Error!"
    Else
        PLP_TS = P
    End If
End Function


Function PHP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2PHP(T, S, P, Range)
    If Range = 0 Then
        PHP_TS = "Error!"
    Else
        PHP_TS = P
    End If
End Function
Function P_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2P(T, S, P, Range)
    If Range = 0 Then
        P_TS = "Error!"
    Else
        P_TS = P
    End If
End Function
Function HLP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2HLP(T, S, H, Range)
    If Range = 0 Then
        HLP_TS = "Error!"
    Else
        HLP_TS = H
    End If
End Function


Function HHP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2HHP(T, S, H, Range)
    If Range = 0 Then
        HHP_TS = "Error!"
    Else
        HHP_TS = H
    End If
End Function
Function H_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2H(T, S, H, Range)
    If Range = 0 Then
        H_TS = "Error!"
    Else
        H_TS = H
    End If
End Function

Function VLP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2VLP(T, S, V, Range)
    If Range = 0 Then
        VLP_TS = "Error!"
    Else
        VLP_TS = V
    End If
End Function

Function VHP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2VHP(T, S, V, Range)
    If Range = 0 Then
        VHP_TS = "Error!"
    Else
        VHP_TS = V
    End If
End Function

Function V_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2V(T, S, V, Range)
    If Range = 0 Then
        V_TS = "Error!"
    Else
        V_TS = V
    End If
End Function
Function X_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute X_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2X(T, S, X, Range)
    If Range = 0 Then
        X_TS = "Error!"
    Else
        X_TS = X
    End If
End Function
Function P_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute P_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2P(T, V, P, Range)
    If Range = 0 Then
        P_TV = "Error!"
    Else
        P_TV = P
    End If
End Function
Function H_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute H_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2H(T, V, H, Range)
    If Range = 0 Then
        H_TV = "Error!"
    Else
        H_TV = H
    End If
End Function
Function S_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute S_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2S(T, V, S, Range)
    If Range = 0 Then
        S_TV = "Error!"
    Else
        S_TV = S
    End If
End Function
Function X_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute X_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2X(T, V, X, Range)
    If Range = 0 Then
        X_TV = "Error!"
    Else
        X_TV = X
    End If
End Function
Function P_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute P_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2P(T, X, P, Range)
    If Range = 0 Then
        P_TX = "Error!"
    Else
        P_TX = P
    End If
End Function
Function H_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute H_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2H(T, X, H, Range)
    If Range = 0 Then
        H_TX = "Error!"
    Else
        H_TX = H
    End If
End Function
Function S_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute S_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2S(T, X, S, Range)
    If Range = 0 Then
        S_TX = "Error!"
    Else
        S_TX = S
    End If
End Function
Function V_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute V_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2V(T, X, V, Range)
    If Range = 0 Then
        V_TX = "Error!"
    Else
        V_TX = V
    End If
End Function


Function P_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute P_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2P(H, S, P, Range)
    If Range = 0 Then
        P_HS = "Error!"
    Else
        P_HS = P
    End If
End Function

Function T_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute T_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的温度T(℃)?"
Rem Attribute T_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2T(H, S, T, Range)
    If Range = 0 Then
        T_HS = "Error!"
    Else
        T_HS = T
    End If
End Function

Function V_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute V_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2V(H, S, V, Range)
    If Range = 0 Then
        V_HS = "Error!"
    Else
        V_HS = V
    End If
End Function

Function X_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute X_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2X(H, S, X, Range)
    If Range = 0 Then
        X_HS = "Error!"
    Else
        X_HS = X
    End If
End Function

Function P_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute P_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2P(H, V, P, Range)
    If Range = 0 Then
        P_HV = "Error!"
    Else
        P_HV = P
    End If
End Function

Function T_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute T_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2T(H, V, T, Range)
    If Range = 0 Then
        T_HV = "Error!"
    Else
        T_HV = T
    End If
End Function

Function S_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute S_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2S(H, V, S, Range)
    If Range = 0 Then
        S_HV = "Error!"
    Else
        S_HV = S
    End If
End Function

Function X_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute X_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2X(H, V, X, Range)
    If Range = 0 Then
        X_HV = "Error!"
    Else
        X_HV = X
    End If
End Function

Function P_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2P(H, X, P, Range)
    If Range = 0 Then
        P_HX = "Error!"
    Else
        P_HX = P
    End If
End Function

Function PLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2PLP(H, X, P, Range)
    If Range = 0 Then
        PLP_HX = "Error!"
    Else
        PLP_HX = P
    End If
End Function

Function PHP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2PHP(H, X, P, Range)
    If Range = 0 Then
        PHP_HX = "Error!"
    Else
        PHP_HX = P
    End If
End Function


Function T_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2T(H, X, T, Range)
    If Range = 0 Then
        T_HX = "Error!"
    Else
        T_HX = T
    End If
End Function

Function TLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2TLP(H, X, T, Range)
    If Range = 0 Then
        TLP_HX = "Error!"
    Else
        TLP_HX = T
    End If
End Function

Function THP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2THP(H, X, T, Range)
    If Range = 0 Then
        THP_HX = "Error!"
    Else
        THP_HX = T
    End If
End Function

Function S_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2S(H, X, S, Range)
    If Range = 0 Then
        S_HX = "Error!"
    Else
        S_HX = S
    End If
End Function

Function SLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2SLP(H, X, S, Range)
    If Range = 0 Then
        SLP_HX = "Error!"
    Else
        SLP_HX = S
    End If
End Function

Function SHP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2SHP(H, X, S, Range)
    If Range = 0 Then
        SHP_HX = "Error!"
    Else
        SHP_HX = S
    End If
End Function

Function V_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2V(H, X, V, Range)
    If Range = 0 Then
        V_HX = "Error!"
    Else
        V_HX = V
    End If
End Function


Function VLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2VLP(H, X, V, Range)
    If Range = 0 Then
        VLP_HX = "Error!"
    Else
        VLP_HX = V
    End If
End Function


Function VHP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2VHP(H, X, V, Range)
    If Range = 0 Then
        VHP_HX = "Error!"
    Else
        VHP_HX = V
    End If
End Function


Function P_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute P_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2P(S, V, P, Range)
    If Range = 0 Then
        P_SV = "Error!"
    Else
        P_SV = P
    End If
End Function

Function T_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute T_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2T(S, V, T, Range)
    If Range = 0 Then
        T_SV = "Error!"
    Else
        T_SV = T
    End If
End Function

Function H_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute H_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2H(S, V, H, Range)
    If Range = 0 Then
        H_SV = "Error!"
    Else
        H_SV = H
    End If
End Function

Function X_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute X_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2X(S, V, X, Range)
    If Range = 0 Then
        X_SV = "Error!"
    Else
        X_SV = X
    End If
End Function

Function P_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2P(S, X, P, Range)
    If Range = 0 Then
        P_SX = "Error!"
    Else
        P_SX = P
    End If
End Function

Function PLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PLP(S, X, P, Range)
    If Range = 0 Then
        PLP_SX = "Error!"
    Else
        PLP_SX = P
    End If
End Function


Function PMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PMP(S, X, P, Range)
    If Range = 0 Then
        PMP_SX = "Error!"
    Else
        PMP_SX = P
    End If
End Function


Function PHP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PHP(S, X, P, Range)
    If Range = 0 Then
        PHP_SX = "Error!"
    Else
        PHP_SX = P
    End If
End Function


Function T_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2T(S, X, T, Range)
    If Range = 0 Then
        T_SX = "Error!"
    Else
        T_SX = T
    End If
End Function

Function TLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2TLP(S, X, T, Range)
    If Range = 0 Then
        TLP_SX = "Error!"
    Else
        TLP_SX = T
    End If
End Function

Function TMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2TMP(S, X, T, Range)
    If Range = 0 Then
        TMP_SX = "Error!"
    Else
        TMP_SX = T
    End If
End Function

Function THP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2THP(S, X, T, Range)
    If Range = 0 Then
        THP_SX = "Error!"
    Else
        THP_SX = T
    End If
End Function

Function H_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2H(S, X, H, Range)
    If Range = 0 Then
        H_SX = "Error!"
    Else
        H_SX = H
    End If
End Function

Function HLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HLP(S, X, H, Range)
    If Range = 0 Then
        HLP_SX = "Error!"
    Else
        HLP_SX = H
    End If
End Function

Function HMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HMP(S, X, H, Range)
    If Range = 0 Then
        HMP_SX = "Error!"
    Else
        HMP_SX = H
    End If
End Function

Function HHP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HHP(S, X, H, Range)
    If Range = 0 Then
        HHP_SX = "Error!"
    Else
        HHP_SX = H
    End If
End Function

Function V_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2V(S, X, V, Range)
    If Range = 0 Then
        V_SX = "Error!"
    Else
        V_SX = V
    End If
End Function

Function VLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VLP(S, X, V, Range)
    If Range = 0 Then
        VLP_SX = "Error!"
    Else
        VLP_SX = V
    End If
End Function

Function VMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VMP(S, X, V, Range)
    If Range = 0 Then
        VMP_SX = "Error!"
    Else
        VMP_SX = V
    End If
End Function

Function VHP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VHP(S, X, V, Range)
    If Range = 0 Then
        VHP_SX = "Error!"
    Else
        VHP_SX = V
    End If
End Function

Function P_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2P(V, X, P, Range)
    If Range = 0 Then
        P_VX = "Error!"
    Else
        P_VX = P
    End If
End Function

Function PLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2PLP(V, X, P, Range)
    If Range = 0 Then
        PLP_VX = "Error!"
    Else
        PLP_VX = P
    End If
End Function

Function PHP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2PHP(V, X, P, Range)
    If Range = 0 Then
        PHP_VX = "Error!"
    Else
        PHP_VX = P
    End If
End Function

Function T_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2T(V, X, T, Range)
    If Range = 0 Then
        T_VX = "Error!"
    Else
        T_VX = T
    End If
End Function

Function TLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2TLP(V, X, T, Range)
    If Range = 0 Then
        TLP_VX = "Error!"
    Else
        TLP_VX = T
    End If
End Function


Function THP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2THP(V, X, T, Range)
    If Range = 0 Then
        THP_VX = "Error!"
    Else
        THP_VX = T
    End If
End Function


Function H_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2H(V, X, H, Range)
    If Range = 0 Then
        H_VX = "Error!"
    Else
        H_VX = H
    End If
End Function

Function HLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2HLP(V, X, H, Range)
    If Range = 0 Then
        HLP_VX = "Error!"
    Else
        HLP_VX = H
    End If
End Function

Function HHP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2HHP(V, X, H, Range)
    If Range = 0 Then
        HHP_VX = "Error!"
    Else
        HHP_VX = H
    End If
End Function

Function S_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2S(V, X, S, Range)
    If Range = 0 Then
        S_VX = "Error!"
    Else
        S_VX = S
    End If
End Function

Function SLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2SLP(V, X, S, Range)
    If Range = 0 Then
        SLP_VX = "Error!"
    Else
        SLP_VX = S
    End If
End Function

Function SHP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2SHP(V, X, S, Range)
    If Range = 0 Then
        SHP_VX = "Error!"
    Else
        SHP_VX = S
    End If
End Function





Rem *************************************************************************************


Function T_P67(ByVal P As Double)
Rem Attribute T_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和温度T(℃)?"
Rem Attribute T_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2T67(P, T, Range)
    If Range = 0 Then
        T_P67 = "Error!"
    Else
        T_P67 = T
    End If
End Function


Function HL_P67(ByVal P As Double)
Rem Attribute Hw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比焓Hw(kJ/kg)?"
Rem Attribute Hw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2HL67(P, H, Range)
    If Range = 0 Then
        HL_P67 = "Error!"
    Else
        HL_P67 = H
    End If
End Function

Function HG_P67(ByVal P As Double)
Rem Attribute Hs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比焓Hs(kJ/kg)?"
Rem Attribute Hs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2HG67(P, H, Range)
    If Range = 0 Then
        HG_P67 = "Error!"
    Else
        HG_P67 = H
    End If
End Function

Function SL_P67(ByVal P As Double)
Rem Attribute Sw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比熵Sw( kJ/(kg.℃) )?"
Rem Attribute Sw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SL67(P, S, Range)
    If Range = 0 Then
        SL_P67 = "Error!"
    Else
        SL_P67 = S
    End If
End Function

Function SG_P67(ByVal P As Double)
Rem Attribute Ss_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比熵Ss( kJ/(kg.℃) )?"
Rem Attribute Ss_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SG67(P, S, Range)
    If Range = 0 Then
        SG_P67 = "Error!"
    Else
        SG_P67 = S
    End If
End Function


Function VL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2VL67(P, V, Range)
    If Range = 0 Then
        VL_P67 = "Error!"
    Else
        VL_P67 = V
    End If
End Function

Function VG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2VG67(P, V, Range)
    If Range = 0 Then
        VG_P67 = "Error!"
    Else
        VG_P67 = V
    End If
End Function


Function CPL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CPL67(P, CP, Range)
    If Range = 0 Then
        CPL_P67 = "Error!"
    Else
        CPL_P67 = CP
    End If
End Function

Function CPG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CPG67(P, CP, Range)
    If Range = 0 Then
        CPG_P67 = "Error!"
    Else
        CPG_P67 = CP
    End If
End Function

Function CVL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CVL67(P, CV, Range)
    If Range = 0 Then
        CVL_P67 = "Error!"
    Else
        CVL_P67 = CV
    End If
End Function

Function CVG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CVG67(P, CV, Range)
    If Range = 0 Then
        CVG_P67 = "Error!"
    Else
        CVG_P67 = CV
    End If
End Function

Function EL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EL67(P, E, Range)
    If Range = 0 Then
        EL_P67 = "Error!"
    Else
        EL_P67 = E
    End If
End Function

Function EG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EG67(P, E, Range)
    If Range = 0 Then
        EG_P67 = "Error!"
    Else
        EG_P67 = E
    End If
End Function

Function SSPL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SSPL67(P, SSP, Range)
    If Range = 0 Then
        SSPL_P67 = "Error!"
    Else
        SSPL_P67 = SSP
    End If
End Function

Function SSPG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SSPG67(P, SSP, Range)
    If Range = 0 Then
        SSPG_P67 = "Error!"
    Else
        SSPG_P67 = SSP
    End If
End Function

Function KSL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2KSL67(P, KS, Range)
    If Range = 0 Then
        KSL_P67 = "Error!"
    Else
        KSL_P67 = KS
    End If
End Function

Function KSG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2KSG67(P, KS, Range)
    If Range = 0 Then
        KSG_P67 = "Error!"
    Else
        KSG_P67 = KS
    End If
End Function


Function ETAL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2ETAL67(P, ETA, Range)
    If Range = 0 Then
        ETAL_P67 = "Error!"
    Else
        ETAL_P67 = ETA
    End If
End Function

Function ETAG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2ETAG67(P, ETA, Range)
    If Range = 0 Then
        ETAG_P67 = "Error!"
    Else
        ETAG_P67 = ETA
    End If
End Function

Function UL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2UL67(P, U, Range)
    If Range = 0 Then
        UL_P67 = "Error!"
    Else
        UL_P67 = U
    End If
End Function

Function UG_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2UG67(P, U, Range)
    If Range = 0 Then
        UG_P67 = "Error!"
    Else
        UG_P67 = U
    End If
End Function

Function RAMDL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2RAMDL67(P, RAMD, Range)
    If Range = 0 Then
        RAMDL_P67 = "Error!"
    Else
        RAMDL_P67 = RAMD
    End If
End Function

Function RAMDG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2RAMDG67(P, RAMD, Range)
    If Range = 0 Then
        RAMDG_P67 = "Error!"
    Else
        RAMDG_P67 = RAMD
    End If
End Function


Function PRNL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2PRNL67(P, PRN, Range)
    If Range = 0 Then
        PRNL_P67 = "Error!"
    Else
        PRNL_P67 = PRN
    End If
End Function

Function PRNG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2PRNG67(P, PRN, Range)
    If Range = 0 Then
        PRNG_P67 = "Error!"
    Else
        PRNG_P67 = PRN
    End If
End Function


Function EPSL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EPSL67(P, eps, Range)
    If Range = 0 Then
        EPSL_P67 = "Error!"
    Else
        EPSL_P67 = eps
    End If
End Function

Function EPSG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EPSG67(P, eps, Range)
    If Range = 0 Then
        EPSG_P67 = "Error!"
    Else
        EPSG_P67 = eps
    End If
End Function

Function NL_P67(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2NL67(P, Lamd, n, Range)
    If Range = 0 Then
        NL_P67 = "Error!"
    Else
        NL_P67 = n
    End If
End Function

Function NG_P67(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2NG67(P, Lamd, n, Range)
    If Range = 0 Then
        NG_P67 = "Error!"
    Else
        NG_P67 = n
    End If
End Function

Function H_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute H_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2H67(P, T, H, Range)
    If Range = 0 Then
        H_PT67 = "Error!"
    Else
        H_PT67 = H
    End If
End Function
Function S_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute S_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2S67(P, T, S, Range)
    If Range = 0 Then
        S_PT67 = "Error!"
    Else
        S_PT67 = S
    End If
End Function
Function V_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute V_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2V67(P, T, V, Range)
    If Range = 0 Then
        V_PT67 = "Error!"
    Else
        V_PT67 = V
    End If
End Function
Function X_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute X_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的干度X(100%)?"
Rem Attribute X_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2X67(P, T, X, Range)
    If Range = 0 Then
        X_PT67 = "Error!"
    Else
        X_PT67 = X
    End If
End Function
Function ETA_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2ETA67(P, T, ETA, Range)
    If Range = 0 Then
        ETA_PT67 = "Error!"
    Else
        ETA_PT67 = ETA
    End If
End Function

Function U_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2U67(P, T, U, Range)
    If Range = 0 Then
        U_PT67 = "Error!"
    Else
        U_PT67 = U
    End If
End Function


Function RAMD_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute RAMD_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的导热系数Ramd( mW/(m.℃) )?"
Rem Attribute RAMD_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求RAMD(导热系数)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2RAMD67(P, T, RAMD, Range)
    If Range = 0 Then
        RAMD_PT67 = "Error!"
    Else
        RAMD_PT67 = RAMD
    End If
End Function

Function CP_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2CP67(P, T, CP, Range)
    If Range = 0 Then
        CP_PT67 = "Error!"
    Else
        CP_PT67 = CP
    End If
End Function

Function CV_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2CV67(P, T, CV, Range)
    If Range = 0 Then
        CV_PT67 = "Error!"
    Else
        CV_PT67 = CV
    End If
End Function

Function E_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2E67(P, T, E, Range)
    If Range = 0 Then
        E_PT67 = "Error!"
    Else
        E_PT67 = E
    End If
End Function
Function KS_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute K_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的绝热指数K(100%)?"
Rem Attribute K_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求K(绝热指数)。
    Dim K As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2KS67(P, T, K, Range)
    If Range = 0 Then
        KS_PT67 = "Error!"
    Else
        KS_PT67 = K
    End If
End Function

Function SSP_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute A_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的音速A (m/s)?"
Rem Attribute A_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求A(音速)。
    Dim a As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2SSP67(P, T, a, Range)
    If Range = 0 Then
        SSP_PT67 = "Error!"
    Else
        SSP_PT67 = a
    End If
End Function

Function PRN_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2PRN67(P, T, PRN, Range)
    If Range = 0 Then
        PRN_PT67 = "Error!"
    Else
        PRN_PT67 = PRN
    End If
End Function

Function EPS_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2EPS67(P, T, eps, Range)
    If Range = 0 Then
        EPS_PT67 = "Error!"
    Else
        EPS_PT67 = eps
    End If
End Function

Function N_PT67(ByVal P As Double, ByVal T As Double, ByVal Lamd As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2N67(P, T, Lamd, n, Range)
    If Range = 0 Then
        N_PT67 = "Error!"
    Else
        N_PT67 = n
    End If
End Function

Function T_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute T_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2T67(P, H, T, Range)
    If Range = 0 Then
        T_PH67 = "Error!"
    Else
        T_PH67 = T
    End If
End Function
Function S_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute S_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2S67(P, H, S, Range)
    If Range = 0 Then
        S_PH67 = "Error!"
    Else
        S_PH67 = S
    End If
End Function
Function V_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute V_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2V67(P, H, V, Range)
    If Range = 0 Then
        V_PH67 = "Error!"
    Else
        V_PH67 = V
    End If
End Function
Function X_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute X_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2X67(P, H, X, Range)
    If Range = 0 Then
        X_PH67 = "Error!"
    Else
        X_PH67 = X
    End If
End Function


Function T_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute T_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的温度T(℃)?"
Rem Attribute T_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2T67(P, S, T, Range)
    If Range = 0 Then
        T_PS67 = "Error!"
    Else
        T_PS67 = T
    End If
End Function
Function H_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute H_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2H67(P, S, H, Range)
    If Range = 0 Then
        H_PS67 = "Error!"
    Else
        H_PS67 = H
    End If
End Function
Function V_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute V_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2V67(P, S, V, Range)
    If Range = 0 Then
        V_PS67 = "Error!"
    Else
        V_PS67 = V
    End If
End Function
Function X_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute X_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2X67(P, S, X, Range)
    If Range = 0 Then
        X_PS67 = "Error!"
    Else
        X_PS67 = X
    End If
End Function


Function T_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute T_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2T67(P, V, T, Range)
    If Range = 0 Then
        T_PV67 = "Error!"
    Else
        T_PV67 = T
    End If
End Function
Function H_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute H_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2H67(P, V, H, Range)
    If Range = 0 Then
        H_PV67 = "Error!"
    Else
        H_PV67 = H
    End If
End Function
Function S_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute S_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2S67(P, V, S, Range)
    If Range = 0 Then
        S_PV67 = "Error!"
    Else
        S_PV67 = S
    End If
End Function
Function X_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute X_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),求对应的干度X(100%)?"
Rem Attribute X_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2X67(P, V, X, Range)
    If Range = 0 Then
        X_PV67 = "Error!"
    Else
        X_PV67 = X
    End If
End Function
Function T_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute T_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2T67(P, X, T, Range)
    If Range = 0 Then
        T_PX67 = "Error!"
    Else
        T_PX67 = T
    End If
End Function
Function H_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute H_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2H67(P, X, H, Range)
    If Range = 0 Then
        H_PX67 = "Error!"
    Else
        H_PX67 = H
    End If
End Function
Function S_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute S_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2S67(P, X, S, Range)
    If Range = 0 Then
        S_PX67 = "Error!"
    Else
        S_PX67 = S
    End If
End Function
Function V_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute V_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2V67(P, X, V, Range)
    If Range = 0 Then
        V_PX67 = "Error!"
    Else
        V_PX67 = V
    End If
End Function


Function P_T67(ByVal T As Double)
Rem Attribute P_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和压力P(MPa)?"
Rem Attribute P_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2P67(T, P, Range)
    If Range = 0 Then
        P_T67 = "Error!"
    Else
        P_T67 = P
    End If
End Function
Function HL_T67(ByVal T As Double)
Rem Attribute Hw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比焓Hw(kJ/kg)?"
Rem Attribute Hw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2HL67(T, H, Range)
    If Range = 0 Then
        HL_T67 = "Error!"
    Else
        HL_T67 = H
    End If
End Function
Function HG_T67(ByVal T As Double)
Rem Attribute Hs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比焓Hs(kJ/kg)?"
Rem Attribute Hs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HS(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2HG67(T, H, Range)
    If Range = 0 Then
        HG_T67 = "Error!"
    Else
        HG_T67 = H
    End If
End Function
Function SG_T67(ByVal T As Double)
Rem Attribute Ss_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比熵Ss( kJ/(kg.℃) )?"
Rem Attribute Ss_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SS(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SG67(T, S, Range)
    If Range = 0 Then
        SG_T67 = "Error!"
    Else
        SG_T67 = S
    End If
End Function
Function SL_T67(ByVal T As Double)
Rem Attribute Sw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比熵Sw( kJ/(kg.℃) )?"
Rem Attribute Sw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SL67(T, S, Range)
    If Range = 0 Then
        SL_T67 = "Error!"
    Else
        SL_T67 = S
    End If
End Function
Function VL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2VL67(T, V, Range)
    If Range = 0 Then
        VL_T67 = "Error!"
    Else
        VL_T67 = V
    End If
End Function
Function VG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2VG67(T, V, Range)
    If Range = 0 Then
        VG_T67 = "Error!"
    Else
        VG_T67 = V
    End If
End Function


Function CPL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CPL67(T, CP, Range)
    If Range = 0 Then
        CPL_T67 = "Error!"
    Else
        CPL_T67 = CP
    End If
End Function
Function CPG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CPG67(T, CP, Range)
    If Range = 0 Then
        CPG_T67 = "Error!"
    Else
        CPG_T67 = CP
    End If
End Function


Function CVL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CVL67(T, CV, Range)
    If Range = 0 Then
        CVL_T67 = "Error!"
    Else
        CVL_T67 = CV
    End If
End Function
Function CVG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CVG67(T, CV, Range)
    If Range = 0 Then
        CVG_T67 = "Error!"
    Else
        CVG_T67 = CV
    End If
End Function

Function EL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EL67(T, E, Range)
    If Range = 0 Then
        EL_T67 = "Error!"
    Else
        EL_T67 = E
    End If
End Function
Function EG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EG67(T, E, Range)
    If Range = 0 Then
        EG_T67 = "Error!"
    Else
        EG_T67 = E
    End If
End Function

Function SSPL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SSPL67(T, SSP, Range)
    If Range = 0 Then
        SSPL_T67 = "Error!"
    Else
        SSPL_T67 = SSP
    End If
End Function
Function SSPG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SSPG67(T, SSP, Range)
    If Range = 0 Then
        SSPG_T67 = "Error!"
    Else
        SSPG_T67 = SSP
    End If
End Function



Function KSL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2KSL67(T, KS, Range)
    If Range = 0 Then
        KSL_T67 = "Error!"
    Else
        KSL_T67 = KS
    End If
End Function
Function KSG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2KSG67(T, KS, Range)
    If Range = 0 Then
        KSG_T67 = "Error!"
    Else
        KSG_T67 = KS
    End If
End Function


Function ETAL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2ETAL67(T, ETA, Range)
    If Range = 0 Then
        ETAL_T67 = "Error!"
    Else
        ETAL_T67 = ETA
    End If
End Function
Function ETAG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2ETAG67(T, ETA, Range)
    If Range = 0 Then
        ETAG_T67 = "Error!"
    Else
        ETAG_T67 = ETA
    End If
End Function

Function UL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2UL67(T, U, Range)
    If Range = 0 Then
        UL_T67 = "Error!"
    Else
        UL_T67 = U
    End If
End Function

Function UG_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2UG67(T, U, Range)
    If Range = 0 Then
        UG_T67 = "Error!"
    Else
        UG_T67 = U
    End If
End Function

Function RAMDL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2RAMDL67(T, RAMD, Range)
    If Range = 0 Then
        RAMDL_T67 = "Error!"
    Else
        RAMDL_T67 = RAMD
    End If
End Function
Function RAMDG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2RAMDG67(T, RAMD, Range)
    If Range = 0 Then
        RAMDG_T67 = "Error!"
    Else
        RAMDG_T67 = RAMD
    End If
End Function




Function PRNL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2PRNL67(T, PRN, Range)
    If Range = 0 Then
        PRNL_T67 = "Error!"
    Else
        PRNL_T67 = PRN
    End If
End Function
Function PRNG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2PRNG67(T, PRN, Range)
    If Range = 0 Then
        PRNG_T67 = "Error!"
    Else
        PRNG_T67 = PRN
    End If
End Function

Function EPSL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EPSL67(T, eps, Range)
    If Range = 0 Then
        EPSL_T67 = "Error!"
    Else
        EPSL_T67 = eps
    End If
End Function
Function EPSG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EPSG67(T, eps, Range)
    If Range = 0 Then
        EPSG_T67 = "Error!"
    Else
        EPSG_T67 = eps
    End If
End Function

Function NL_T67(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2NL67(T, Lamd, n, Range)
    If Range = 0 Then
        NL_T67 = "Error!"
    Else
        NL_T67 = n
    End If
End Function

Function NG_T67(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2NG67(T, Lamd, n, Range)
    If Range = 0 Then
        NG_T67 = "Error!"
    Else
        NG_T67 = n
    End If
End Function

Function SurfT_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SurfT As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SURFT67(T, SurfT, Range)
    If Range = 0 Then
        SurfT_T67 = "Error!"
    Else
        SurfT_T67 = SurfT
    End If
End Function

Function P_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2P67(T, H, P, Range)
    If Range = 0 Then
        P_TH67 = "Error!"
    Else
        P_TH67 = P
    End If
End Function

Function PLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2PLP67(T, H, P, Range)
    If Range = 0 Then
        PLP_TH67 = "Error!"
    Else
        PLP_TH67 = P
    End If
End Function



Function PHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2PHP67(T, H, P, Range)
    If Range = 0 Then
        PHP_TH67 = "Error!"
    Else
        PHP_TH67 = P
    End If
End Function

Function S_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2S67(T, H, S, Range)
    If Range = 0 Then
        S_TH67 = "Error!"
    Else
        S_TH67 = S
    End If
End Function

Function SLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2SLP67(T, H, S, Range)
    If Range = 0 Then
        SLP_TH67 = "Error!"
    Else
        SLP_TH67 = S
    End If
End Function

Function SHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2SHP67(T, H, S, Range)
    If Range = 0 Then
        SHP_TH67 = "Error!"
    Else
        SHP_TH67 = S
    End If
End Function


Function V_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2V67(T, H, V, Range)
    If Range = 0 Then
        V_TH67 = "Error!"
    Else
        V_TH67 = V
    End If
End Function


Function VLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2VLP67(T, H, V, Range)
    If Range = 0 Then
        VLP_TH67 = "Error!"
    Else
        VLP_TH67 = V
    End If
End Function


Function VHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2VHP67(T, H, V, Range)
    If Range = 0 Then
        VHP_TH67 = "Error!"
    Else
        VHP_TH67 = V
    End If
End Function

Function XLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2XLP67(T, H, X, Range)
    If Range = 0 Then
        XLP_TH67 = "Error!"
    Else
        XLP_TH67 = X
    End If
End Function
Function XHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2XHP67(T, H, X, Range)
    If Range = 0 Then
        XHP_TH67 = "Error!"
    Else
        XHP_TH67 = X
    End If
End Function
Function X_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2X67(T, H, X, Range)
    If Range = 0 Then
        X_TH67 = "Error!"
    Else
        X_TH67 = X
    End If
End Function


Function PLP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2PLP67(T, S, P, Range)
    If Range = 0 Then
        PLP_TS67 = "Error!"
    Else
        PLP_TS67 = P
    End If
End Function


Function PHP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2PHP67(T, S, P, Range)
    If Range = 0 Then
        PHP_TS67 = "Error!"
    Else
        PHP_TS67 = P
    End If
End Function
Function P_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2P67(T, S, P, Range)
    If Range = 0 Then
        P_TS67 = "Error!"
    Else
        P_TS67 = P
    End If
End Function
Function HLP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2HLP67(T, S, H, Range)
    If Range = 0 Then
        HLP_TS67 = "Error!"
    Else
        HLP_TS67 = H
    End If
End Function


Function HHP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2HHP67(T, S, H, Range)
    If Range = 0 Then
        HHP_TS67 = "Error!"
    Else
        HHP_TS67 = H
    End If
End Function
Function H_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2H67(T, S, H, Range)
    If Range = 0 Then
        H_TS67 = "Error!"
    Else
        H_TS67 = H
    End If
End Function

Function VLP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2VLP67(T, S, V, Range)
    If Range = 0 Then
        VLP_TS67 = "Error!"
    Else
        VLP_TS67 = V
    End If
End Function

Function VHP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2VHP67(T, S, V, Range)
    If Range = 0 Then
        VHP_TS67 = "Error!"
    Else
        VHP_TS67 = V
    End If
End Function

Function V_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2V67(T, S, V, Range)
    If Range = 0 Then
        V_TS67 = "Error!"
    Else
        V_TS67 = V
    End If
End Function
Function X_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute X_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2X67(T, S, X, Range)
    If Range = 0 Then
        X_TS67 = "Error!"
    Else
        X_TS67 = X
    End If
End Function
Function P_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute P_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2P67(T, V, P, Range)
    If Range = 0 Then
        P_TV67 = "Error!"
    Else
        P_TV67 = P
    End If
End Function
Function H_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute H_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2H67(T, V, H, Range)
    If Range = 0 Then
        H_TV67 = "Error!"
    Else
        H_TV67 = H
    End If
End Function
Function S_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute S_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2S67(T, V, S, Range)
    If Range = 0 Then
        S_TV67 = "Error!"
    Else
        S_TV67 = S
    End If
End Function
Function X_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute X_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2X67(T, V, X, Range)
    If Range = 0 Then
        X_TV67 = "Error!"
    Else
        X_TV67 = X
    End If
End Function
Function P_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute P_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2P67(T, X, P, Range)
    If Range = 0 Then
        P_TX67 = "Error!"
    Else
        P_TX67 = P
    End If
End Function
Function H_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute H_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2H67(T, X, H, Range)
    If Range = 0 Then
        H_TX67 = "Error!"
    Else
        H_TX67 = H
    End If
End Function
Function S_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute S_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2S67(T, X, S, Range)
    If Range = 0 Then
        S_TX67 = "Error!"
    Else
        S_TX67 = S
    End If
End Function
Function V_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute V_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2V67(T, X, V, Range)
    If Range = 0 Then
        V_TX67 = "Error!"
    Else
        V_TX67 = V
    End If
End Function


Function P_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute P_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2P67(H, S, P, Range)
    If Range = 0 Then
        P_HS67 = "Error!"
    Else
        P_HS67 = P
    End If
End Function

Function T_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute T_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的温度T(℃)?"
Rem Attribute T_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2T67(H, S, T, Range)
    If Range = 0 Then
        T_HS67 = "Error!"
    Else
        T_HS67 = T
    End If
End Function

Function V_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute V_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2V67(H, S, V, Range)
    If Range = 0 Then
        V_HS67 = "Error!"
    Else
        V_HS67 = V
    End If
End Function

Function X_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute X_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2X67(H, S, X, Range)
    If Range = 0 Then
        X_HS67 = "Error!"
    Else
        X_HS67 = X
    End If
End Function

Function P_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute P_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2P67(H, V, P, Range)
    If Range = 0 Then
        P_HV67 = "Error!"
    Else
        P_HV67 = P
    End If
End Function

Function T_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute T_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2T67(H, V, T, Range)
    If Range = 0 Then
        T_HV67 = "Error!"
    Else
        T_HV67 = T
    End If
End Function

Function S_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute S_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2S67(H, V, S, Range)
    If Range = 0 Then
        S_HV67 = "Error!"
    Else
        S_HV67 = S
    End If
End Function

Function X_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute X_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2X67(H, V, X, Range)
    If Range = 0 Then
        X_HV67 = "Error!"
    Else
        X_HV67 = X
    End If
End Function

Function P_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2P67(H, X, P, Range)
    If Range = 0 Then
        P_HX67 = "Error!"
    Else
        P_HX67 = P
    End If
End Function

Function PLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2PLP67(H, X, P, Range)
    If Range = 0 Then
        PLP_HX67 = "Error!"
    Else
        PLP_HX67 = P
    End If
End Function

Function PHP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2PHP67(H, X, P, Range)
    If Range = 0 Then
        PHP_HX67 = "Error!"
    Else
        PHP_HX67 = P
    End If
End Function


Function T_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2T67(H, X, T, Range)
    If Range = 0 Then
        T_HX67 = "Error!"
    Else
        T_HX67 = T
    End If
End Function

Function TLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2TLP67(H, X, T, Range)
    If Range = 0 Then
        TLP_HX67 = "Error!"
    Else
        TLP_HX67 = T
    End If
End Function

Function THP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2THP67(H, X, T, Range)
    If Range = 0 Then
        THP_HX67 = "Error!"
    Else
        THP_HX67 = T
    End If
End Function

Function S_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2S67(H, X, S, Range)
    If Range = 0 Then
        S_HX67 = "Error!"
    Else
        S_HX67 = S
    End If
End Function

Function SLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2SLP67(H, X, S, Range)
    If Range = 0 Then
        SLP_HX67 = "Error!"
    Else
        SLP_HX67 = S
    End If
End Function

Function SHP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2SHP67(H, X, S, Range)
    If Range = 0 Then
        SHP_HX67 = "Error!"
    Else
        SHP_HX67 = S
    End If
End Function

Function V_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2V67(H, X, V, Range)
    If Range = 0 Then
        V_HX67 = "Error!"
    Else
        V_HX67 = V
    End If
End Function


Function VLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2VLP67(H, X, V, Range)
    If Range = 0 Then
        VLP_HX67 = "Error!"
    Else
        VLP_HX67 = V
    End If
End Function


Function VHP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2VHP67(H, X, V, Range)
    If Range = 0 Then
        VHP_HX67 = "Error!"
    Else
        VHP_HX67 = V
    End If
End Function


Function P_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute P_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2P67(S, V, P, Range)
    If Range = 0 Then
        P_SV67 = "Error!"
    Else
        P_SV67 = P
    End If
End Function

Function T_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute T_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2T67(S, V, T, Range)
    If Range = 0 Then
        T_SV67 = "Error!"
    Else
        T_SV67 = T
    End If
End Function

Function H_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute H_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2H67(S, V, H, Range)
    If Range = 0 Then
        H_SV67 = "Error!"
    Else
        H_SV67 = H
    End If
End Function

Function X_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute X_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2X67(S, V, X, Range)
    If Range = 0 Then
        X_SV67 = "Error!"
    Else
        X_SV67 = X
    End If
End Function

Function P_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2P67(S, X, P, Range)
    If Range = 0 Then
        P_SX67 = "Error!"
    Else
        P_SX67 = P
    End If
End Function

Function PLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PLP67(S, X, P, Range)
    If Range = 0 Then
        PLP_SX67 = "Error!"
    Else
        PLP_SX67 = P
    End If
End Function


Function PMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PMP67(S, X, P, Range)
    If Range = 0 Then
        PMP_SX67 = "Error!"
    Else
        PMP_SX67 = P
    End If
End Function


Function PHP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PHP67(S, X, P, Range)
    If Range = 0 Then
        PHP_SX67 = "Error!"
    Else
        PHP_SX67 = P
    End If
End Function


Function T_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2T67(S, X, T, Range)
    If Range = 0 Then
        T_SX67 = "Error!"
    Else
        T_SX67 = T
    End If
End Function

Function TLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2TLP67(S, X, T, Range)
    If Range = 0 Then
        TLP_SX67 = "Error!"
    Else
        TLP_SX67 = T
    End If
End Function

Function TMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2TMP67(S, X, T, Range)
    If Range = 0 Then
        TMP_SX67 = "Error!"
    Else
        TMP_SX67 = T
    End If
End Function

Function THP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2THP67(S, X, T, Range)
    If Range = 0 Then
        THP_SX67 = "Error!"
    Else
        THP_SX67 = T
    End If
End Function

Function H_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2H67(S, X, H, Range)
    If Range = 0 Then
        H_SX67 = "Error!"
    Else
        H_SX67 = H
    End If
End Function

Function HLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HLP67(S, X, H, Range)
    If Range = 0 Then
        HLP_SX67 = "Error!"
    Else
        HLP_SX67 = H
    End If
End Function

Function HMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HMP67(S, X, H, Range)
    If Range = 0 Then
        HMP_SX67 = "Error!"
    Else
        HMP_SX67 = H
    End If
End Function

Function HHP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HHP67(S, X, H, Range)
    If Range = 0 Then
        HHP_SX67 = "Error!"
    Else
        HHP_SX67 = H
    End If
End Function

Function V_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2V67(S, X, V, Range)
    If Range = 0 Then
        V_SX67 = "Error!"
    Else
        V_SX67 = V
    End If
End Function

Function VLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VLP67(S, X, V, Range)
    If Range = 0 Then
        VLP_SX67 = "Error!"
    Else
        VLP_SX67 = V
    End If
End Function

Function VMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VMP67(S, X, V, Range)
    If Range = 0 Then
        VMP_SX67 = "Error!"
    Else
        VMP_SX67 = V
    End If
End Function

Function VHP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VHP67(S, X, V, Range)
    If Range = 0 Then
        VHP_SX67 = "Error!"
    Else
        VHP_SX67 = V
    End If
End Function

Function P_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2P67(V, X, P, Range)
    If Range = 0 Then
        P_VX67 = "Error!"
    Else
        P_VX67 = P
    End If
End Function

Function PLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2PLP67(V, X, P, Range)
    If Range = 0 Then
        PLP_VX67 = "Error!"
    Else
        PLP_VX67 = P
    End If
End Function

Function PHP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2PHP67(V, X, P, Range)
    If Range = 0 Then
        PHP_VX67 = "Error!"
    Else
        PHP_VX67 = P
    End If
End Function

Function T_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2T67(V, X, T, Range)
    If Range = 0 Then
        T_VX67 = "Error!"
    Else
        T_VX67 = T
    End If
End Function

Function TLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2TLP67(V, X, T, Range)
    If Range = 0 Then
        TLP_VX67 = "Error!"
    Else
        TLP_VX67 = T
    End If
End Function


Function THP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2THP67(V, X, T, Range)
    If Range = 0 Then
        THP_VX67 = "Error!"
    Else
        THP_VX67 = T
    End If
End Function


Function H_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2H67(V, X, H, Range)
    If Range = 0 Then
        H_VX67 = "Error!"
    Else
        H_VX67 = H
    End If
End Function

Function HLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2HLP67(V, X, H, Range)
    If Range = 0 Then
        HLP_VX67 = "Error!"
    Else
        HLP_VX67 = H
    End If
End Function

Function HHP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2HHP67(V, X, H, Range)
    If Range = 0 Then
        HHP_VX67 = "Error!"
    Else
        HHP_VX67 = H
    End If
End Function

Function S_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2S67(V, X, S, Range)
    If Range = 0 Then
        S_VX67 = "Error!"
    Else
        S_VX67 = S
    End If
End Function

Function SLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2SLP67(V, X, S, Range)
    If Range = 0 Then
        SLP_VX67 = "Error!"
    Else
        SLP_VX67 = S
    End If
End Function

Function SHP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2SHP67(V, X, S, Range)
    If Range = 0 Then
        SHP_VX67 = "Error!"
    Else
        SHP_VX67 = S
    End If
End Function



Rem *************************************************************************************

Function T_P97(ByVal P As Double)
Rem Attribute T_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和温度T(℃)?"
Rem Attribute T_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2T97(P, T, Range)
    If Range = 0 Then
        T_P97 = "Error!"
    Else
        T_P97 = T
    End If
End Function


Function HL_P97(ByVal P As Double)
Rem Attribute Hw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比焓Hw(kJ/kg)?"
Rem Attribute Hw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2HL97(P, H, Range)
    If Range = 0 Then
        HL_P97 = "Error!"
    Else
        HL_P97 = H
    End If
End Function

Function HG_P97(ByVal P As Double)
Rem Attribute Hs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比焓Hs(kJ/kg)?"
Rem Attribute Hs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2HG97(P, H, Range)
    If Range = 0 Then
        HG_P97 = "Error!"
    Else
        HG_P97 = H
    End If
End Function

Function SL_P97(ByVal P As Double)
Rem Attribute Sw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比熵Sw( kJ/(kg.℃) )?"
Rem Attribute Sw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SL97(P, S, Range)
    If Range = 0 Then
        SL_P97 = "Error!"
    Else
        SL_P97 = S
    End If
End Function

Function SG_P97(ByVal P As Double)
Rem Attribute Ss_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比熵Ss( kJ/(kg.℃) )?"
Rem Attribute Ss_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SG97(P, S, Range)
    If Range = 0 Then
        SG_P97 = "Error!"
    Else
        SG_P97 = S
    End If
End Function


Function VL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2VL97(P, V, Range)
    If Range = 0 Then
        VL_P97 = "Error!"
    Else
        VL_P97 = V
    End If
End Function

Function VG_P97(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2VG97(P, V, Range)
    If Range = 0 Then
        VG_P97 = "Error!"
    Else
        VG_P97 = V
    End If
End Function


Function CpL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CPL97(P, CP, Range)
    If Range = 0 Then
        CpL_P97 = "Error!"
    Else
        CpL_P97 = CP
    End If
End Function

Function CpG_P97(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CPG97(P, CP, Range)
    If Range = 0 Then
        CpG_P97 = "Error!"
    Else
        CpG_P97 = CP
    End If
End Function

Function CvL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CVL97(P, CV, Range)
    If Range = 0 Then
        CvL_P97 = "Error!"
    Else
        CvL_P97 = CV
    End If
End Function

Function CvG_P97(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2CVG97(P, CV, Range)
    If Range = 0 Then
        CvG_P97 = "Error!"
    Else
        CvG_P97 = CV
    End If
End Function


Function EL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EL97(P, E, Range)
    If Range = 0 Then
        EL_P97 = "Error!"
    Else
        EL_P97 = E
    End If
End Function


Function EG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EG97(P, E, Range)
    If Range = 0 Then
        EG_P97 = "Error!"
    Else
        EG_P97 = E
    End If
End Function


Function SSpL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SSPL97(P, SSP, Range)
    If Range = 0 Then
        SSpL_P97 = "Error!"
    Else
        SSpL_P97 = SSP
    End If
End Function

Function SSpG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2SSPG97(P, SSP, Range)
    If Range = 0 Then
        SSpG_P97 = "Error!"
    Else
        SSpG_P97 = SSP
    End If
End Function

Function KsL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2KSL97(P, KS, Range)
    If Range = 0 Then
        KsL_P97 = "Error!"
    Else
        KsL_P97 = KS
    End If
End Function

Function KsG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2KSG97(P, KS, Range)
    If Range = 0 Then
        KsG_P97 = "Error!"
    Else
        KsG_P97 = KS
    End If
End Function

Function EtaL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2ETAL97(P, ETA, Range)
    If Range = 0 Then
        EtaL_P97 = "Error!"
    Else
        EtaL_P97 = ETA
    End If
End Function


Function EtaG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2ETAG97(P, ETA, Range)
    If Range = 0 Then
        EtaG_P97 = "Error!"
    Else
        EtaG_P97 = ETA
    End If
End Function

Function UL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2UL97(P, U, Range)
    If Range = 0 Then
        UL_P97 = "Error!"
    Else
        UL_P97 = U
    End If
End Function

Function UG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2UG97(P, U, Range)
    If Range = 0 Then
        UG_P97 = "Error!"
    Else
        UG_P97 = U
    End If
End Function

Function RamdL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2RAMDL97(P, RAMD, Range)
    If Range = 0 Then
        RamdL_P97 = "Error!"
    Else
        RamdL_P97 = RAMD
    End If
End Function


Function RamdG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2RAMDG97(P, RAMD, Range)
    If Range = 0 Then
        RamdG_P97 = "Error!"
    Else
        RamdG_P97 = RAMD
    End If
End Function

Function EpsL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EPSL97(P, eps, Range)
    If Range = 0 Then
        EpsL_P97 = "Error!"
    Else
        EpsL_P97 = eps
    End If
End Function

Function EpsG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2EPSG97(P, eps, Range)
    If Range = 0 Then
        EpsG_P97 = "Error!"
    Else
        EpsG_P97 = eps
    End If
End Function

Function PrnL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2PRNL97(P, PRN, Range)
    If Range = 0 Then
        PrnL_P97 = "Error!"
    Else
        PrnL_P97 = PRN
    End If
End Function

Function PrnG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2PRNG97(P, PRN, Range)
    If Range = 0 Then
        PrnG_P97 = "Error!"
    Else
        PrnG_P97 = PRN
    End If
End Function

Function NL_P97(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2NL97(P, Lamd, n, Range)
    If Range = 0 Then
        NL_P97 = "Error!"
    Else
        NL_P97 = n
    End If
End Function

Function NG_P97(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(MPa),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    Call P2NG97(P, Lamd, n, Range)
    If Range = 0 Then
        NG_P97 = "Error!"
    Else
        NG_P97 = n
    End If
End Function

Function H_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute H_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2H97(P, T, H, Range)
    If Range = 0 Then
        H_PT97 = "Error!"
    Else
        H_PT97 = H
    End If
End Function
Function S_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute S_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2S97(P, T, S, Range)
    If Range = 0 Then
        S_PT97 = "Error!"
    Else
        S_PT97 = S
    End If
End Function
Function V_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute V_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2V97(P, T, V, Range)
    If Range = 0 Then
        V_PT97 = "Error!"
    Else
        V_PT97 = V
    End If
End Function
Function X_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute X_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的干度X(100%)?"
Rem Attribute X_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2X97(P, T, X, Range)
    If Range = 0 Then
        X_PT97 = "Error!"
    Else
        X_PT97 = X
    End If
End Function


Function Cp_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2CP97(P, T, CP, Range)
    If Range = 0 Then
        Cp_PT97 = "Error!"
    Else
        Cp_PT97 = CP
    End If
End Function


Function Cv_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2CV97(P, T, CV, Range)
    If Range = 0 Then
        Cv_PT97 = "Error!"
    Else
        Cv_PT97 = CV
    End If
End Function

Function E_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2E97(P, T, E, Range)
    If Range = 0 Then
        E_PT97 = "Error!"
    Else
        E_PT97 = E
    End If
End Function


Function SSp_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2SSP97(P, T, SSP, Range)
    If Range = 0 Then
        SSp_PT97 = "Error!"
    Else
        SSp_PT97 = SSP
    End If
End Function


Function Ks_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的定压质量比热容CP( kJ/(kg.℃) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2KS97(P, T, KS, Range)
    If Range = 0 Then
        Ks_PT97 = "Error!"
    Else
        Ks_PT97 = KS
    End If
End Function


Function Eta_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2ETA97(P, T, ETA, Range)
    If Range = 0 Then
        Eta_PT97 = "Error!"
    Else
        Eta_PT97 = ETA
    End If
End Function

Function U_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2U97(P, T, U, Range)
    If Range = 0 Then
        U_PT97 = "Error!"
    Else
        U_PT97 = U
    End If
End Function


Function Ramd_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute RAMD_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的导热系数Ramd( mW/(m.℃) )?"
Rem Attribute RAMD_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求RAMD(导热系数)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2RAMD97(P, T, RAMD, Range)
    If Range = 0 Then
        Ramd_PT97 = "Error!"
    Else
        Ramd_PT97 = RAMD
    End If
End Function


Function PRN_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2PRN97(P, T, PRN, Range)
    If Range = 0 Then
        PRN_PT97 = "Error!"
    Else
        PRN_PT97 = PRN
    End If
End Function

Function Eps_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2EPS97(P, T, eps, Range)
    If Range = 0 Then
        Eps_PT97 = "Error!"
    Else
        Eps_PT97 = eps
    End If
End Function

Function N_PT97(ByVal P As Double, ByVal T As Double, ByVal Lamd As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(MPa)和温度T(℃),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    Call PT2N97(P, T, Lamd, n, Range)
    If Range = 0 Then
        N_PT97 = "Error!"
    Else
        N_PT97 = n
    End If
End Function

Function T_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute T_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2T97(P, H, T, Range)
    If Range = 0 Then
        T_PH97 = "Error!"
    Else
        T_PH97 = T
    End If
End Function
Function S_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute S_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2S97(P, H, S, Range)
    If Range = 0 Then
        S_PH97 = "Error!"
    Else
        S_PH97 = S
    End If
End Function
Function V_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute V_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2V97(P, H, V, Range)
    If Range = 0 Then
        V_PH97 = "Error!"
    Else
        V_PH97 = V
    End If
End Function
Function X_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute X_PH.VB_Description = "已知压力P(MPa)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    Call PH2X97(P, H, X, Range)
    If Range = 0 Then
        X_PH97 = "Error!"
    Else
        X_PH97 = X
    End If
End Function


Function T_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute T_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的温度T(℃)?"
Rem Attribute T_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2T97(P, S, T, Range)
    If Range = 0 Then
        T_PS97 = "Error!"
    Else
        T_PS97 = T
    End If
End Function
Function H_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute H_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2H97(P, S, H, Range)
    If Range = 0 Then
        H_PS97 = "Error!"
    Else
        H_PS97 = H
    End If
End Function
Function V_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute V_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2V97(P, S, V, Range)
    If Range = 0 Then
        V_PS97 = "Error!"
    Else
        V_PS97 = V
    End If
End Function
Function X_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute X_PS.VB_Description = "已知压力P(MPa)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    Call PS2X97(P, S, X, Range)
    If Range = 0 Then
        X_PS97 = "Error!"
    Else
        X_PS97 = X
    End If
End Function


Function T_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute T_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2T97(P, V, T, Range)
    If Range = 0 Then
        T_PV97 = "Error!"
    Else
        T_PV97 = T
    End If
End Function
Function H_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute H_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2H97(P, V, H, Range)
    If Range = 0 Then
        H_PV97 = "Error!"
    Else
        H_PV97 = H
    End If
End Function
Function S_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute S_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2S97(P, V, S, Range)
    If Range = 0 Then
        S_PV97 = "Error!"
    Else
        S_PV97 = S
    End If
End Function
Function X_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute X_PV.VB_Description = "已知压力P(MPa)和比容V(m^3/kg),求对应的干度X(100%)?"
Rem Attribute X_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    Call PV2X97(P, V, X, Range)
    If Range = 0 Then
        X_PV97 = "Error!"
    Else
        X_PV97 = X
    End If
End Function
Function T_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute T_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2T97(P, X, T, Range)
    If Range = 0 Then
        T_PX97 = "Error!"
    Else
        T_PX97 = T
    End If
End Function
Function H_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute H_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2H97(P, X, H, Range)
    If Range = 0 Then
        H_PX97 = "Error!"
    Else
        H_PX97 = H
    End If
End Function
Function S_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute S_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2S97(P, X, S, Range)
    If Range = 0 Then
        S_PX97 = "Error!"
    Else
        S_PX97 = S
    End If
End Function
Function V_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute V_PX.VB_Description = "已知压力P(MPa)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    Call PX2V97(P, X, V, Range)
    If Range = 0 Then
        V_PX97 = "Error!"
    Else
        V_PX97 = V
    End If
End Function


Function P_T97(ByVal T As Double)
Rem Attribute P_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和压力P(MPa)?"
Rem Attribute P_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2P97(T, P, Range)
    If Range = 0 Then
        P_T97 = "Error!"
    Else
        P_T97 = P
    End If
End Function
Function HL_T97(ByVal T As Double)
Rem Attribute Hw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比焓Hw(kJ/kg)?"
Rem Attribute Hw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2HL97(T, H, Range)
    If Range = 0 Then
        HL_T97 = "Error!"
    Else
        HL_T97 = H
    End If
End Function
Function HG_T97(ByVal T As Double)
Rem Attribute Hs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比焓Hs(kJ/kg)?"
Rem Attribute Hs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HS(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2HG97(T, H, Range)
    If Range = 0 Then
        HG_T97 = "Error!"
    Else
        HG_T97 = H
    End If
End Function
Function SL_T97(ByVal T As Double)
Rem Attribute Ss_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比熵Ss( kJ/(kg.℃) )?"
Rem Attribute Ss_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SS(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SL97(T, S, Range)
    If Range = 0 Then
        SL_T97 = "Error!"
    Else
        SL_T97 = S
    End If
End Function
Function SG_T97(ByVal T As Double)
Rem Attribute Sw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比熵Sw( kJ/(kg.℃) )?"
Rem Attribute Sw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SG97(T, S, Range)
    If Range = 0 Then
        SG_T97 = "Error!"
    Else
        SG_T97 = S
    End If
End Function
Function VL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2VL97(T, V, Range)
    If Range = 0 Then
        VL_T97 = "Error!"
    Else
        VL_T97 = V
    End If
End Function
Function VG_T97(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和蒸汽比容Vs(m^3/kg)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2VG97(T, V, Range)
    If Range = 0 Then
        VG_T97 = "Error!"
    Else
        VG_T97 = V
    End If
End Function


Function CpL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CPL97(T, CP, Range)
    If Range = 0 Then
        CpL_T97 = "Error!"
    Else
        CpL_T97 = CP
    End If
End Function


Function CpG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CPG97(T, CP, Range)
    If Range = 0 Then
        CpG_T97 = "Error!"
    Else
        CpG_T97 = CP
    End If
End Function


Function CvL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CVL97(T, CV, Range)
    If Range = 0 Then
        CvL_T97 = "Error!"
    Else
        CvL_T97 = CV
    End If
End Function



Function CvG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2CVG97(T, CV, Range)
    If Range = 0 Then
        CvG_T97 = "Error!"
    Else
        CvG_T97 = CV
    End If
End Function

Function EL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EL97(T, E, Range)
    If Range = 0 Then
        EL_T97 = "Error!"
    Else
        EL_T97 = E
    End If
End Function


Function EG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EG97(T, E, Range)
    If Range = 0 Then
        EG_T97 = "Error!"
    Else
        EG_T97 = E
    End If
End Function


Function SSpL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SSPL97(T, SSP, Range)
    If Range = 0 Then
        SSpL_T97 = "Error!"
    Else
        SSpL_T97 = SSP
    End If
End Function


Function SSpG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SSPG97(T, SSP, Range)
    If Range = 0 Then
        SSpG_T97 = "Error!"
    Else
        SSpG_T97 = SSP
    End If
End Function

Function KsL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2KSL97(T, KS, Range)
    If Range = 0 Then
        KsL_T97 = "Error!"
    Else
        KsL_T97 = KS
    End If
End Function

Function KsG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2KSG97(T, KS, Range)
    If Range = 0 Then
        KsG_T97 = "Error!"
    Else
        KsG_T97 = KS
    End If
End Function

Function EtaL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2ETAL97(T, ETA, Range)
    If Range = 0 Then
        EtaL_T97 = "Error!"
    Else
        EtaL_T97 = ETA
    End If
End Function



Function EtaG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2ETAG97(T, ETA, Range)
    If Range = 0 Then
        EtaG_T97 = "Error!"
    Else
        EtaG_T97 = ETA
    End If
End Function


Function UL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2UL97(T, U, Range)
    If Range = 0 Then
        UL_T97 = "Error!"
    Else
        UL_T97 = U
    End If
End Function

Function UG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2UG97(T, U, Range)
    If Range = 0 Then
        UG_T97 = "Error!"
    Else
        UG_T97 = U
    End If
End Function

Function RamdL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2RAMDL97(T, RAMD, Range)
    If Range = 0 Then
        RamdL_T97 = "Error!"
    Else
        RamdL_T97 = RAMD
    End If
End Function


Function RamdG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2RAMDG97(T, RAMD, Range)
    If Range = 0 Then
        RamdG_T97 = "Error!"
    Else
        RamdG_T97 = RAMD
    End If
End Function


Function EpsL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EPSL97(T, eps, Range)
    If Range = 0 Then
        EpsL_T97 = "Error!"
    Else
        EpsL_T97 = eps
    End If
End Function

Function EpsG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2EPSG97(T, eps, Range)
    If Range = 0 Then
        EpsG_T97 = "Error!"
    Else
        EpsG_T97 = eps
    End If
End Function

Function PrnL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2PRNL97(T, PRN, Range)
    If Range = 0 Then
        PrnL_T97 = "Error!"
    Else
        PrnL_T97 = PRN
    End If
End Function

Function PrnG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2PRNG97(T, PRN, Range)
    If Range = 0 Then
        PrnG_T97 = "Error!"
    Else
        PrnG_T97 = PRN
    End If
End Function

Function NL_T97(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2NL97(T, Lamd, n, Range)
    If Range = 0 Then
        NL_T97 = "Error!"
    Else
        NL_T97 = n
    End If
End Function

Function NG_T97(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2NG97(T, Lamd, n, Range)
    If Range = 0 Then
        NG_T97 = "Error!"
    Else
        NG_T97 = n
    End If
End Function

Function SurfT_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(℃),\r\n求对应的饱和水比容Vw(m^3/kg)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SurfT As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    Call T2SURFT97(T, SurfT, Range)
    If Range = 0 Then
        SurfT_T97 = "Error!"
    Else
        SurfT_T97 = SurfT
    End If
End Function

Function P_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2P97(T, H, P, Range)
    If Range = 0 Then
        P_TH97 = "Error!"
    Else
        P_TH97 = P
    End If
End Function

Function PLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2PLP97(T, H, P, Range)
    If Range = 0 Then
        PLP_TH97 = "Error!"
    Else
        PLP_TH97 = P
    End If
End Function


Function PHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2PHP97(T, H, P, Range)
    If Range = 0 Then
        PHP_TH97 = "Error!"
    Else
        PHP_TH97 = P
    End If
End Function

Function S_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2S97(T, H, S, Range)
    If Range = 0 Then
        S_TH97 = "Error!"
    Else
        S_TH97 = S
    End If
End Function
Function SLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2SLP97(T, H, S, Range)
    If Range = 0 Then
        SLP_TH97 = "Error!"
    Else
        SLP_TH97 = S
    End If
End Function



Function SHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2SHP97(T, H, S, Range)
    If Range = 0 Then
        SHP_TH97 = "Error!"
    Else
        SHP_TH97 = S
    End If
End Function

Function V_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2V97(T, H, V, Range)
    If Range = 0 Then
        V_TH97 = "Error!"
    Else
        V_TH97 = V
    End If
End Function
Function VLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2VLP97(T, H, V, Range)
    If Range = 0 Then
        VLP_TH97 = "Error!"
    Else
        VLP_TH97 = V
    End If
End Function
Function VHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2VHP97(T, H, V, Range)
    If Range = 0 Then
        VHP_TH97 = "Error!"
    Else
        VHP_TH97 = V
    End If
End Function

Function XLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2XLP97(T, H, X, Range)
    If Range = 0 Then
        XLP_TH97 = "Error!"
    Else
        XLP_TH97 = X
    End If
End Function

Function XHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2XHP97(T, H, X, Range)
    If Range = 0 Then
        XHP_TH97 = "Error!"
    Else
        XHP_TH97 = X
    End If
End Function

Function X_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(℃)和比焓H(kJ/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    Call TH2X97(T, H, X, Range)
    If Range = 0 Then
        X_TH97 = "Error!"
    Else
        X_TH97 = X
    End If
End Function


Function P_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2P97(T, S, P, Range)
    If Range = 0 Then
        P_TS97 = "Error!"
    Else
        P_TS97 = P
    End If
End Function

Function PLP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2PLP97(T, S, P, Range)
    If Range = 0 Then
        PLP_TS97 = "Error!"
    Else
        PLP_TS97 = P
    End If
End Function


Function PHP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2PHP97(T, S, P, Range)
    If Range = 0 Then
        PHP_TS97 = "Error!"
    Else
        PHP_TS97 = P
    End If
End Function



Function H_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2H97(T, S, H, Range)
    If Range = 0 Then
        H_TS97 = "Error!"
    Else
        H_TS97 = H
    End If
End Function


Function HLP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2HLP97(T, S, H, Range)
    If Range = 0 Then
        HLP_TS97 = "Error!"
    Else
        HLP_TS97 = H
    End If
End Function


Function HHP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2HHP97(T, S, H, Range)
    If Range = 0 Then
        HHP_TS97 = "Error!"
    Else
        HHP_TS97 = H
    End If
End Function




Function V_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2V97(T, S, V, Range)
    If Range = 0 Then
        V_TS97 = "Error!"
    Else
        V_TS97 = V
    End If
End Function

Function VLP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2VLP97(T, S, V, Range)
    If Range = 0 Then
        VLP_TS97 = "Error!"
    Else
        VLP_TS97 = V
    End If
End Function


Function VHP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2VHP97(T, S, V, Range)
    If Range = 0 Then
        VHP_TS97 = "Error!"
    Else
        VHP_TS97 = V
    End If
End Function


Function X_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute X_TS.VB_Description = "已知温度T(℃)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    Call TS2X97(T, S, X, Range)
    If Range = 0 Then
        X_TS97 = "Error!"
    Else
        X_TS97 = X
    End If
End Function
Function P_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute P_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2P97(T, V, P, Range)
    If Range = 0 Then
        P_TV97 = "Error!"
    Else
        P_TV97 = P
    End If
End Function
Function H_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute H_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2H97(T, V, H, Range)
    If Range = 0 Then
        H_TV97 = "Error!"
    Else
        H_TV97 = H
    End If
End Function
Function S_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute S_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2S97(T, V, S, Range)
    If Range = 0 Then
        S_TV97 = "Error!"
    Else
        S_TV97 = S
    End If
End Function
Function X_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute X_TV.VB_Description = "已知温度T(℃)和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    Call TV2X97(T, V, X, Range)
    If Range = 0 Then
        X_TV97 = "Error!"
    Else
        X_TV97 = X
    End If
End Function
Function P_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute P_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2P97(T, X, P, Range)
    If Range = 0 Then
        P_TX97 = "Error!"
    Else
        P_TX97 = P
    End If
End Function
Function H_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute H_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2H97(T, X, H, Range)
    If Range = 0 Then
        H_TX97 = "Error!"
    Else
        H_TX97 = H
    End If
End Function
Function S_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute S_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2S97(T, X, S, Range)
    If Range = 0 Then
        S_TX97 = "Error!"
    Else
        S_TX97 = S
    End If
End Function
Function V_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute V_TX.VB_Description = "已知温度T(℃)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    Call TX2V97(T, X, V, Range)
    If Range = 0 Then
        V_TX97 = "Error!"
    Else
        V_TX97 = V
    End If
End Function


Function P_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute P_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2P97(H, S, P, Range)
    If Range = 0 Then
        P_HS97 = "Error!"
    Else
        P_HS97 = P
    End If
End Function

Function T_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute T_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的温度T(℃)?"
Rem Attribute T_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2T97(H, S, T, Range)
    If Range = 0 Then
        T_HS97 = "Error!"
    Else
        T_HS97 = T
    End If
End Function

Function V_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute V_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2V97(H, S, V, Range)
    If Range = 0 Then
        V_HS97 = "Error!"
    Else
        V_HS97 = V
    End If
End Function

Function X_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute X_HS.VB_Description = "已知比焓H(kJ/kg)和比熵S( kJ/(kg.℃) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    Call HS2X97(H, S, X, Range)
    If Range = 0 Then
        X_HS97 = "Error!"
    Else
        X_HS97 = X
    End If
End Function

Function P_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute P_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2P97(H, V, P, Range)
    If Range = 0 Then
        P_HV97 = "Error!"
    Else
        P_HV97 = P
    End If
End Function

Function T_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute T_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2T97(H, V, T, Range)
    If Range = 0 Then
        T_HV97 = "Error!"
    Else
        T_HV97 = T
    End If
End Function

Function S_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute S_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2S97(H, V, S, Range)
    If Range = 0 Then
        S_HV97 = "Error!"
    Else
        S_HV97 = S
    End If
End Function

Function X_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute X_HV.VB_Description = "已知比焓H(kJ/kg)和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    Call HV2X97(H, V, X, Range)
    If Range = 0 Then
        X_HV97 = "Error!"
    Else
        X_HV97 = X
    End If
End Function

Function P_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2P97(H, X, P, Range)
    If Range = 0 Then
        P_HX97 = "Error!"
    Else
        P_HX97 = P
    End If
End Function

Function PLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2PLP97(H, X, P, Range)
    If Range = 0 Then
        PLP_HX97 = "Error!"
    Else
        PLP_HX97 = P
    End If
End Function


Function PHP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2PHP97(H, X, P, Range)
    If Range = 0 Then
        PHP_HX97 = "Error!"
    Else
        PHP_HX97 = P
    End If
End Function


Function T_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2T97(H, X, T, Range)
    If Range = 0 Then
        T_HX97 = "Error!"
    Else
        T_HX97 = T
    End If
End Function


Function TLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2TLP97(H, X, T, Range)
    If Range = 0 Then
        TLP_HX97 = "Error!"
    Else
        TLP_HX97 = T
    End If
End Function


Function THP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2THP97(H, X, T, Range)
    If Range = 0 Then
        THP_HX97 = "Error!"
    Else
        THP_HX97 = T
    End If
End Function

Function S_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2S97(H, X, S, Range)
    If Range = 0 Then
        S_HX97 = "Error!"
    Else
        S_HX97 = S
    End If
End Function

Function SLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2SLP97(H, X, S, Range)
    If Range = 0 Then
        SLP_HX97 = "Error!"
    Else
        SLP_HX97 = S
    End If
End Function

Function SHP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2SHP97(H, X, S, Range)
    If Range = 0 Then
        SHP_HX97 = "Error!"
    Else
        SHP_HX97 = S
    End If
End Function

Function V_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2V97(H, X, V, Range)
    If Range = 0 Then
        V_HX97 = "Error!"
    Else
        V_HX97 = V
    End If
End Function

Function VLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2VLP97(H, X, V, Range)
    If Range = 0 Then
        VLP_HX97 = "Error!"
    Else
        VLP_HX97 = V
    End If
End Function

Function VHP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(kJ/kg)和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    Call HX2VHP97(H, X, V, Range)
    If Range = 0 Then
        VHP_HX97 = "Error!"
    Else
        VHP_HX97 = V
    End If
End Function


Function P_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute P_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2P97(S, V, P, Range)
    If Range = 0 Then
        P_SV97 = "Error!"
    Else
        P_SV97 = P
    End If
End Function

Function T_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute T_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的温度T(℃)?"
Rem Attribute T_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2T97(S, V, T, Range)
    If Range = 0 Then
        T_SV97 = "Error!"
    Else
        T_SV97 = T
    End If
End Function

Function H_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute H_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2H97(S, V, H, Range)
    If Range = 0 Then
        H_SV97 = "Error!"
    Else
        H_SV97 = H
    End If
End Function

Function X_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute X_SV.VB_Description = "已知比熵S( kJ/(kg.℃) )和比容V(m^3/kg),\r\n求对应的干度X(100%)?"
Rem Attribute X_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    Call SV2X97(S, V, X, Range)
    If Range = 0 Then
        X_SV97 = "Error!"
    Else
        X_SV97 = X
    End If
End Function

Function P_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2P97(S, X, P, Range)
    If Range = 0 Then
        P_SX97 = "Error!"
    Else
        P_SX97 = P
    End If
End Function


Function PLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PLP97(S, X, P, Range)
    If Range = 0 Then
        PLP_SX97 = "Error!"
    Else
        PLP_SX97 = P
    End If
End Function

Function PMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PMP97(S, X, P, Range)
    If Range = 0 Then
        PMP_SX97 = "Error!"
    Else
        PMP_SX97 = P
    End If
End Function

Function PHP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2PHP97(S, X, P, Range)
    If Range = 0 Then
        PHP_SX97 = "Error!"
    Else
        PHP_SX97 = P
    End If
End Function
Function T_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2T97(S, X, T, Range)
    If Range = 0 Then
        T_SX97 = "Error!"
    Else
        T_SX97 = T
    End If
End Function

Function TLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2TLP97(S, X, T, Range)
    If Range = 0 Then
        TLP_SX97 = "Error!"
    Else
        TLP_SX97 = T
    End If
End Function

Function TMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2TMP97(S, X, T, Range)
    If Range = 0 Then
        TMP_SX97 = "Error!"
    Else
        TMP_SX97 = T
    End If
End Function

Function THP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2THP97(S, X, T, Range)
    If Range = 0 Then
        THP_SX97 = "Error!"
    Else
        THP_SX97 = T
    End If
End Function

Function H_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2H97(S, X, H, Range)
    If Range = 0 Then
        H_SX97 = "Error!"
    Else
        H_SX97 = H
    End If
End Function

Function HLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HLP97(S, X, H, Range)
    If Range = 0 Then
        HLP_SX97 = "Error!"
    Else
        HLP_SX97 = H
    End If
End Function

Function HMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HMP97(S, X, H, Range)
    If Range = 0 Then
        HMP_SX97 = "Error!"
    Else
        HMP_SX97 = H
    End If
End Function

Function HHP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2HHP97(S, X, H, Range)
    If Range = 0 Then
        HHP_SX97 = "Error!"
    Else
        HHP_SX97 = H
    End If
End Function

Function V_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2V97(S, X, V, Range)
    If Range = 0 Then
        V_SX97 = "Error!"
    Else
        V_SX97 = V
    End If
End Function

Function VLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VLP97(S, X, V, Range)
    If Range = 0 Then
        VLP_SX97 = "Error!"
    Else
        VLP_SX97 = V
    End If
End Function

Function VMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VMP97(S, X, V, Range)
    If Range = 0 Then
        VMP_SX97 = "Error!"
    Else
        VMP_SX97 = V
    End If
End Function

Function VHP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( kJ/(kg.℃) )和干度X(100%),\r\n求对应的比容V(m^3/kg)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    Call SX2VHP97(S, X, V, Range)
    If Range = 0 Then
        VHP_SX97 = "Error!"
    Else
        VHP_SX97 = V
    End If
End Function

Function P_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2P97(V, X, P, Range)
    If Range = 0 Then
        P_VX97 = "Error!"
    Else
        P_VX97 = P
    End If
End Function
Function PLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2PLP97(V, X, P, Range)
    If Range = 0 Then
        PLP_VX97 = "Error!"
    Else
        PLP_VX97 = P
    End If
End Function
Function PHP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的压力P(MPa)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2PHP97(V, X, P, Range)
    If Range = 0 Then
        PHP_VX97 = "Error!"
    Else
        PHP_VX97 = P
    End If
End Function

Function T_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2T97(V, X, T, Range)
    If Range = 0 Then
        T_VX97 = "Error!"
    Else
        T_VX97 = T
    End If
End Function

Function TLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2TLP97(V, X, T, Range)
    If Range = 0 Then
        TLP_VX97 = "Error!"
    Else
        TLP_VX97 = T
    End If
End Function

Function THP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的温度T(℃)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2THP97(V, X, T, Range)
    If Range = 0 Then
        THP_VX97 = "Error!"
    Else
        THP_VX97 = T
    End If
End Function

Function H_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2H97(V, X, H, Range)
    If Range = 0 Then
        H_VX97 = "Error!"
    Else
        H_VX97 = H
    End If
End Function

Function HLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2HLP97(V, X, H, Range)
    If Range = 0 Then
        HLP_VX97 = "Error!"
    Else
        HLP_VX97 = H
    End If
End Function

Function HHP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比焓H(kJ/kg)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2HHP97(V, X, H, Range)
    If Range = 0 Then
        HHP_VX97 = "Error!"
    Else
        HHP_VX97 = H
    End If
End Function

Function S_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2S97(V, X, S, Range)
    If Range = 0 Then
        S_VX97 = "Error!"
    Else
        S_VX97 = S
    End If
End Function

Function SLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2SLP97(V, X, S, Range)
    If Range = 0 Then
        SLP_VX97 = "Error!"
    Else
        SLP_VX97 = S
    End If
End Function

Function SHP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(m^3/kg)和干度X(100%),\r\n求对应的比熵S( kJ/(kg.℃) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    Call VX2SHP97(V, X, S, Range)
    If Range = 0 Then
        SHP_VX97 = "Error!"
    Else
        SHP_VX97 = S
    End If
End Function

Rem 两点线性插值
Rem function INT2DXX(ByVal X1 As Double, ByVal X2 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal x As Double) As Double
Rem Attribute IN2DX_Y.VB_Description = "根据两点线性插值"
Rem Attribute IN2DX_Y.VB_ProcData.VB_Invoke_Func = " \n16"
Rem    Dim y As Double
Rem    Call INST2DXX(X1, X2, Y1, Y2, x, y)
Rem    INT2DXX = y
Rem End Function


Rem function INT2DXY(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal x As Double) As Double
Rem Attribute P2_XY.VB_Description = "根据两点线性插值"
Rem Attribute P2_XY.VB_ProcData.VB_Invoke_Func = " \n16"
Rem    Dim y As Double
Rem    Call INST2DXY(X1, Y1, X2, Y2, x, y)
Rem    INT2DXY = y
Rem End Function


Rem ================================================================================================================================
Rem ================================================================================================================================
Rem $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位
Rem $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位
Rem $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位
Rem $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位
Rem $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位
Rem $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位 $$英制单位
Rem ================================================================================================================================
Rem ================================================================================================================================
Rem 英制单位水蒸汽函数 在原来的函数名前加US_作为标记


Function US_T_P(ByVal P As Double)
Rem Attribute T_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和温度T(F)?"
Rem Attribute T_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double
Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2T(P, T, Range)
    If Range = 0 Then
        US_T_P = "Error!"
    Else
        US_T_P = T * 1.8 + 32
    End If
End Function


Function US_HL_P(ByVal P As Double)
Rem Attribute Hw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比焓Hw(Btu/lbm)?"
Rem Attribute Hw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2HL(P, H, Range)
    If Range = 0 Then
        US_HL_P = "Error!"
    Else
        US_HL_P = H / 2.326
    End If
End Function

Function US_HG_P(ByVal P As Double)
Rem Attribute Hs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比焓Hs(Btu/lbm)?"
Rem Attribute Hs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2HG(P, H, Range)
    If Range = 0 Then
        US_HG_P = "Error!"
    Else
        US_HG_P = H / 2.326
    End If
End Function

Function US_SL_P(ByVal P As Double)
Rem Attribute Sw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比熵Sw( (Btu/lbmR) )?"
Rem Attribute Sw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SL(P, S, Range)
    If Range = 0 Then
        US_SL_P = "Error!"
    Else
        US_SL_P = S / 4.1868
    End If
End Function

Function US_SG_P(ByVal P As Double)
Rem Attribute Ss_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比熵Ss( (Btu/lbmR) )?"
Rem Attribute Ss_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SG(P, S, Range)
    If Range = 0 Then
        US_SG_P = "Error!"
    Else
        US_SG_P = S / 4.1868
    End If
End Function


Function US_VL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2VL(P, V, Range)
    If Range = 0 Then
        US_VL_P = "Error!"
    Else
        US_VL_P = V / 0.062428
    
End If
End Function

Function US_VG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2VG(P, V, Range)
    If Range = 0 Then
        US_VG_P = "Error!"
    Else
        US_VG_P = V / 0.062428
    
End If
End Function


Function US_CPL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CPL(P, CP, Range)
    If Range = 0 Then
        US_CPL_P = "Error!"
    Else
        US_CPL_P = CP
    End If
End Function

Function US_CPG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CPG(P, CP, Range)
    If Range = 0 Then
        US_CPG_P = "Error!"
    Else
        US_CPG_P = CP
    End If
End Function

Function US_CVL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CVL(P, CV, Range)
    If Range = 0 Then
        US_CVL_P = "Error!"
    Else
        US_CVL_P = CV
    End If
End Function

Function US_CVG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CVG(P, CV, Range)
    If Range = 0 Then
        US_CVG_P = "Error!"
    Else
        US_CVG_P = CV
    End If
End Function

Function US_EL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EL(P, E, Range)
    If Range = 0 Then
        US_EL_P = "Error!"
    Else
        US_EL_P = E
    End If
End Function

Function US_EG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EG(P, E, Range)
    If Range = 0 Then
        US_EG_P = "Error!"
    Else
        US_EG_P = E
    End If
End Function

Function US_SSPL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SSPL(P, SSP, Range)
    If Range = 0 Then
        US_SSPL_P = "Error!"
    Else
        US_SSPL_P = SSP
    End If
End Function

Function US_SSPG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SSPG(P, SSP, Range)
    If Range = 0 Then
        US_SSPG_P = "Error!"
    Else
        US_SSPG_P = SSP
    End If
End Function

Function US_KSL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2KSL(P, KS, Range)
    If Range = 0 Then
        US_KSL_P = "Error!"
    Else
        US_KSL_P = KS
    End If
End Function

Function US_KSG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2KSG(P, KS, Range)
    If Range = 0 Then
        US_KSG_P = "Error!"
    Else
        US_KSG_P = KS
    End If
End Function


Function US_ETAL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2ETAL(P, ETA, Range)
    If Range = 0 Then
        US_ETAL_P = "Error!"
    Else
        US_ETAL_P = ETA
    End If
End Function

Function US_ETAG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2ETAG(P, ETA, Range)
    If Range = 0 Then
        US_ETAG_P = "Error!"
    Else
        US_ETAG_P = ETA
    End If
End Function

Function US_UL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2UL(P, U, Range)
    If Range = 0 Then
        US_UL_P = "Error!"
    Else
        US_UL_P = U
    End If
End Function

Function US_UG_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2UG(P, U, Range)
    If Range = 0 Then
        US_UG_P = "Error!"
    Else
        US_UG_P = U
    End If
End Function

Function US_RAMDL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2RAMDL(P, RAMD, Range)
    If Range = 0 Then
        US_RAMDL_P = "Error!"
    Else
        US_RAMDL_P = RAMD
    End If
End Function

Function US_RAMDG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2RAMDG(P, RAMD, Range)
    If Range = 0 Then
        US_RAMDG_P = "Error!"
    Else
        US_RAMDG_P = RAMD
    End If
End Function


Function US_PRNL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2PRNL(P, PRN, Range)
    If Range = 0 Then
        US_PRNL_P = "Error!"
    Else
        US_PRNL_P = PRN
    End If
End Function

Function US_PRNG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2PRNG(P, PRN, Range)
    If Range = 0 Then
        US_PRNG_P = "Error!"
    Else
        US_PRNG_P = PRN
    End If
End Function


Function US_EPSL_P(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EPSL(P, eps, Range)
    If Range = 0 Then
        US_EPSL_P = "Error!"
    Else
        US_EPSL_P = eps
    End If
End Function

Function US_EPSG_P(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EPSG(P, eps, Range)
    If Range = 0 Then
        US_EPSG_P = "Error!"
    Else
        US_EPSG_P = eps
    End If
End Function

Function US_NL_P(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2NL(P, Lamd, n, Range)
    If Range = 0 Then
        US_NL_P = "Error!"
    Else
        US_NL_P = n
    End If
End Function

Function US_NG_P(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2NG(P, Lamd, n, Range)
    If Range = 0 Then
        US_NG_P = "Error!"
    Else
        US_NG_P = n
    End If
End Function

Function US_H_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute H_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2H(P, T, H, Range)
    If Range = 0 Then
        US_H_PT = "Error!"
    Else
        US_H_PT = H / 2.326
    End If
End Function
Function US_S_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute S_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2S(P, T, S, Range)
    If Range = 0 Then
        US_S_PT = "Error!"
    Else
        US_S_PT = S / 4.1868
    End If
End Function
Function US_V_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute V_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2V(P, T, V, Range)
    If Range = 0 Then
        US_V_PT = "Error!"
    Else
        US_V_PT = V / 0.062428
    
End If
End Function
Function US_X_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute X_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的干度X(100%)?"
Rem Attribute X_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2X(P, T, X, Range)
    If Range = 0 Then
        US_X_PT = "Error!"
    Else
        US_X_PT = X
    End If
End Function
Function US_ETA_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2ETA(P, T, ETA, Range)
    If Range = 0 Then
        US_ETA_PT = "Error!"
    Else
        US_ETA_PT = ETA
    End If
End Function

Function US_U_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2U(P, T, U, Range)
    If Range = 0 Then
        US_U_PT = "Error!"
    Else
        US_U_PT = U
    End If
End Function


Function US_RAMD_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute RAMD_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的导热系数Ramd( mW/(m.℃) )?"
Rem Attribute RAMD_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求RAMD(导热系数)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2RAMD(P, T, RAMD, Range)
    If Range = 0 Then
        US_RAMD_PT = "Error!"
    Else
        US_RAMD_PT = RAMD
    End If
End Function

Function US_CP_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2CP(P, T, CP, Range)
    If Range = 0 Then
        US_CP_PT = "Error!"
    Else
        US_CP_PT = CP
    End If
End Function

Function US_CV_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2CV(P, T, CV, Range)
    If Range = 0 Then
        US_CV_PT = "Error!"
    Else
        US_CV_PT = CV
    End If
End Function

Function US_E_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2E(P, T, E, Range)
    If Range = 0 Then
        US_E_PT = "Error!"
    Else
        US_E_PT = E
    End If
End Function
Function US_KS_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute K_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的绝热指数K(100%)?"
Rem Attribute K_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求K(绝热指数)。
    Dim K As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2KS(P, T, K, Range)
    If Range = 0 Then
        US_KS_PT = "Error!"
    Else
        US_KS_PT = K
    End If
End Function

Function US_SSP_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute A_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的音速A (m/s)?"
Rem Attribute A_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求A(音速)。
    Dim a As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2SSP(P, T, a, Range)
    If Range = 0 Then
        US_SSP_PT = "Error!"
    Else
        US_SSP_PT = a
    End If
End Function

Function US_PRN_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2PRN(P, T, PRN, Range)
    If Range = 0 Then
        US_PRN_PT = "Error!"
    Else
        US_PRN_PT = PRN
    End If
End Function

Function US_EPS_PT(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2EPS(P, T, eps, Range)
    If Range = 0 Then
        US_EPS_PT = "Error!"
    Else
        US_EPS_PT = eps
    End If
End Function

Function US_N_PT(ByVal P As Double, ByVal T As Double, ByVal Lamd As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2N(P, T, Lamd, n, Range)
    If Range = 0 Then
        US_N_PT = "Error!"
    Else
        US_N_PT = n
    End If
End Function

Function US_T_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute T_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2T(P, H, T, Range)
    If Range = 0 Then
        US_T_PH = "Error!"
    Else
        US_T_PH = T * 1.8 + 32
    End If
End Function
Function US_S_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute S_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2S(P, H, S, Range)
    If Range = 0 Then
        US_S_PH = "Error!"
    Else
        US_S_PH = S / 4.1868
    End If
End Function
Function US_v_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute V_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2V(P, H, V, Range)
    If Range = 0 Then
        US_v_PH = "Error!"
    Else
        US_v_PH = V / 0.062428
    
End If
End Function
Function US_X_PH(ByVal P As Double, ByVal H As Double)
Rem Attribute X_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2X(P, H, X, Range)
    If Range = 0 Then
        US_X_PH = "Error!"
    Else
        US_X_PH = X
    End If
End Function


Function US_T_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute T_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的温度T(F)?"
Rem Attribute T_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2T(P, S, T, Range)
    If Range = 0 Then
        US_T_PS = "Error!"
    Else
        US_T_PS = T * 1.8 + 32
    End If
End Function
Function US_H_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute H_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2H(P, S, H, Range)
    If Range = 0 Then
        US_H_PS = "Error!"
    Else
        US_H_PS = H / 2.326
    End If
End Function
Function US_V_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute V_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2V(P, S, V, Range)
    If Range = 0 Then
        US_V_PS = "Error!"
    Else
        US_V_PS = V / 0.062428
    
End If
End Function
Function US_X_PS(ByVal P As Double, ByVal S As Double)
Rem Attribute X_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2X(P, S, X, Range)
    If Range = 0 Then
        US_X_PS = "Error!"
    Else
        US_X_PS = X
    End If
End Function


Function US_T_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute T_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2T(P, V, T, Range)
    If Range = 0 Then
        US_T_PV = "Error!"
    Else
        US_T_PV = T * 1.8 + 32
    End If
End Function
Function US_H_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute H_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2H(P, V, H, Range)
    If Range = 0 Then
        US_H_PV = "Error!"
    Else
        US_H_PV = H / 2.326
    End If
End Function
Function US_S_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute S_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2S(P, V, S, Range)
    If Range = 0 Then
        US_S_PV = "Error!"
    Else
        US_S_PV = S / 4.1868
    End If
End Function
Function US_X_PV(ByVal P As Double, ByVal V As Double)
Rem Attribute X_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),求对应的干度X(100%)?"
Rem Attribute X_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2X(P, V, X, Range)
    If Range = 0 Then
        US_X_PV = "Error!"
    Else
        US_X_PV = X
    End If
End Function
Function US_T_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute T_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2T(P, X, T, Range)
    If Range = 0 Then
        US_T_PX = "Error!"
    Else
        US_T_PX = T * 1.8 + 32
    End If
End Function
Function US_H_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute H_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2H(P, X, H, Range)
    If Range = 0 Then
        US_H_PX = "Error!"
    Else
        US_H_PX = H / 2.326
    End If
End Function
Function US_S_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute S_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2S(P, X, S, Range)
    If Range = 0 Then
        US_S_PX = "Error!"
    Else
        US_S_PX = S / 4.1868
    End If
End Function
Function US_V_PX(ByVal P As Double, ByVal X As Double)
Rem Attribute V_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2V(P, X, V, Range)
    If Range = 0 Then
        US_V_PX = "Error!"
    Else
        US_V_PX = V / 0.062428
    
End If
End Function


Function US_P_T(ByVal T As Double)
Rem Attribute P_T.VB_Description = "已知温度T(F),\r\n求对应的饱和压力P(Psi)?"
Rem Attribute P_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2P(T, P, Range)
    If Range = 0 Then
        US_P_T = "Error!"
    Else
        US_P_T = P * 10 / 0.068948
    
End If
End Function
Function US_HL_T(ByVal T As Double)
Rem Attribute Hw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比焓Hw(Btu/lbm)?"
Rem Attribute Hw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2HL(T, H, Range)
    If Range = 0 Then
        US_HL_T = "Error!"
    Else
        US_HL_T = H / 2.326
    End If
End Function
Function US_HG_T(ByVal T As Double)
Rem Attribute Hs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比焓Hs(Btu/lbm)?"
Rem Attribute Hs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HS(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2HG(T, H, Range)
    If Range = 0 Then
        US_HG_T = "Error!"
    Else
        US_HG_T = H / 2.326
    End If
End Function
Function US_SG_T(ByVal T As Double)
Rem Attribute Ss_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比熵Ss( (Btu/lbmR) )?"
Rem Attribute Ss_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SS(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SG(T, S, Range)
    If Range = 0 Then
        US_SG_T = "Error!"
    Else
        US_SG_T = S / 4.1868
    End If
End Function
Function US_SL_T(ByVal T As Double)
Rem Attribute Sw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比熵Sw( (Btu/lbmR) )?"
Rem Attribute Sw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SL(T, S, Range)
    If Range = 0 Then
        US_SL_T = "Error!"
    Else
        US_SL_T = S / 4.1868
    End If
End Function
Function US_VL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2VL(T, V, Range)
    If Range = 0 Then
        US_VL_T = "Error!"
    Else
        US_VL_T = V / 0.062428
    
End If
End Function
Function US_VG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2VG(T, V, Range)
    If Range = 0 Then
        US_VG_T = "Error!"
    Else
        US_VG_T = V / 0.062428
    
End If
End Function


Function US_CPL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CPL(T, CP, Range)
    If Range = 0 Then
        US_CPL_T = "Error!"
    Else
        US_CPL_T = CP
    End If
End Function
Function US_CPG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CPG(T, CP, Range)
    If Range = 0 Then
        US_CPG_T = "Error!"
    Else
        US_CPG_T = CP
    End If
End Function


Function US_CVL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CVL(T, CV, Range)
    If Range = 0 Then
        US_CVL_T = "Error!"
    Else
        US_CVL_T = CV
    End If
End Function
Function US_CVG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CVG(T, CV, Range)
    If Range = 0 Then
        US_CVG_T = "Error!"
    Else
        US_CVG_T = CV
    End If
End Function

Function US_EL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EL(T, E, Range)
    If Range = 0 Then
        US_EL_T = "Error!"
    Else
        US_EL_T = E
    End If
End Function
Function US_EG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EG(T, E, Range)
    If Range = 0 Then
        US_EG_T = "Error!"
    Else
        US_EG_T = E
    End If
End Function

Function US_SSPL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SSPL(T, SSP, Range)
    If Range = 0 Then
        US_SSPL_T = "Error!"
    Else
        US_SSPL_T = SSP
    End If
End Function
Function US_SSPG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SSPG(T, SSP, Range)
    If Range = 0 Then
        US_SSPG_T = "Error!"
    Else
        US_SSPG_T = SSP
    End If
End Function



Function US_KSL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2KSL(T, KS, Range)
    If Range = 0 Then
        US_KSL_T = "Error!"
    Else
        US_KSL_T = KS
    End If
End Function
Function US_KSG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2KSG(T, KS, Range)
    If Range = 0 Then
        US_KSG_T = "Error!"
    Else
        US_KSG_T = KS
    End If
End Function


Function US_ETAL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2ETAL(T, ETA, Range)
    If Range = 0 Then
        US_ETAL_T = "Error!"
    Else
        US_ETAL_T = ETA
    End If
End Function
Function US_ETAG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2ETAG(T, ETA, Range)
    If Range = 0 Then
        US_ETAG_T = "Error!"
    Else
        US_ETAG_T = ETA
    End If
End Function

Function US_UL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2UL(T, U, Range)
    If Range = 0 Then
        US_UL_T = "Error!"
    Else
        US_UL_T = U
    End If
End Function

Function US_UG_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2UG(T, U, Range)
    If Range = 0 Then
        US_UG_T = "Error!"
    Else
        US_UG_T = U
    End If
End Function

Function US_RAMDL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2RAMDL(T, RAMD, Range)
    If Range = 0 Then
        US_RAMDL_T = "Error!"
    Else
        US_RAMDL_T = RAMD
    End If
End Function
Function US_RAMDG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2RAMDG(T, RAMD, Range)
    If Range = 0 Then
        US_RAMDG_T = "Error!"
    Else
        US_RAMDG_T = RAMD
    End If
End Function




Function US_PRNL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2PRNL(T, PRN, Range)
    If Range = 0 Then
        US_PRNL_T = "Error!"
    Else
        US_PRNL_T = PRN
    End If
End Function
Function US_PRNG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2PRNG(T, PRN, Range)
    If Range = 0 Then
        US_PRNG_T = "Error!"
    Else
        US_PRNG_T = PRN
    End If
End Function

Function US_EPSL_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EPSL(T, eps, Range)
    If Range = 0 Then
        US_EPSL_T = "Error!"
    Else
        US_EPSL_T = eps
    End If
End Function
Function US_EPSG_T(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EPSG(T, eps, Range)
    If Range = 0 Then
        US_EPSG_T = "Error!"
    Else
        US_EPSG_T = eps
    End If
End Function

Function US_NL_T(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2NL(T, Lamd, n, Range)
    If Range = 0 Then
        US_NL_T = "Error!"
    Else
        US_NL_T = n
    End If
End Function

Function US_NG_T(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2NG(T, Lamd, n, Range)
    If Range = 0 Then
        US_NG_T = "Error!"
    Else
        US_NG_T = n
    End If
End Function

Function US_SurfT_T(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SurfT As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SURFT(T, SurfT, Range)
    If Range = 0 Then
        US_SurfT_T = "Error!"
    Else
        US_SurfT_T = SurfT
    End If
End Function

Function US_P_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2P(T, H, P, Range)
    If Range = 0 Then
        US_P_TH = "Error!"
    Else
        US_P_TH = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2PLP(T, H, P, Range)
    If Range = 0 Then
        US_PLP_TH = "Error!"
    Else
        US_PLP_TH = P * 10 / 0.068948
    
End If
End Function



Function US_PHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2PHP(T, H, P, Range)
    If Range = 0 Then
        US_PHP_TH = "Error!"
    Else
        US_PHP_TH = P * 10 / 0.068948
    
End If
End Function

Function US_S_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2S(T, H, S, Range)
    If Range = 0 Then
        US_S_TH = "Error!"
    Else
        US_S_TH = S / 4.1868
    End If
End Function

Function US_SLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2SLP(T, H, S, Range)
    If Range = 0 Then
        US_SLP_TH = "Error!"
    Else
        US_SLP_TH = S / 4.1868
    End If
End Function

Function US_SHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2SHP(T, H, S, Range)
    If Range = 0 Then
        US_SHP_TH = "Error!"
    Else
        US_SHP_TH = S / 4.1868
    End If
End Function


Function US_V_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2V(T, H, V, Range)
    If Range = 0 Then
        US_V_TH = "Error!"
    Else
        US_V_TH = V / 0.062428
    
End If
End Function


Function US_VLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2VLP(T, H, V, Range)
    If Range = 0 Then
        US_VLP_TH = "Error!"
    Else
        US_VLP_TH = V / 0.062428
    
End If
End Function


Function US_VHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2VHP(T, H, V, Range)
    If Range = 0 Then
        US_VHP_TH = "Error!"
    Else
        US_VHP_TH = V / 0.062428
    
End If
End Function

Function US_XLP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2XLP(T, H, X, Range)
    If Range = 0 Then
        US_XLP_TH = "Error!"
    Else
        US_XLP_TH = X
    End If
End Function
Function US_XHP_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2XHP(T, H, X, Range)
    If Range = 0 Then
        US_XHP_TH = "Error!"
    Else
        US_XHP_TH = X
    End If
End Function
Function US_X_TH(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2X(T, H, X, Range)
    If Range = 0 Then
        US_X_TH = "Error!"
    Else
        US_X_TH = X
    End If
End Function


Function US_PLP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2PLP(T, S, P, Range)
    If Range = 0 Then
        US_PLP_TS = "Error!"
    Else
        US_PLP_TS = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2PHP(T, S, P, Range)
    If Range = 0 Then
        US_PHP_TS = "Error!"
    Else
        US_PHP_TS = P * 10 / 0.068948
    
End If
End Function
Function US_P_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2P(T, S, P, Range)
    If Range = 0 Then
        US_P_TS = "Error!"
    Else
        US_P_TS = P * 10 / 0.068948
    
End If
End Function
Function US_HLP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2HLP(T, S, H, Range)
    If Range = 0 Then
        US_HLP_TS = "Error!"
    Else
        US_HLP_TS = H / 2.326
    End If
End Function


Function US_HHP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2HHP(T, S, H, Range)
    If Range = 0 Then
        US_HHP_TS = "Error!"
    Else
        US_HHP_TS = H / 2.326
    End If
End Function
Function US_H_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2H(T, S, H, Range)
    If Range = 0 Then
        US_H_TS = "Error!"
    Else
        US_H_TS = H / 2.326
    End If
End Function

Function US_VLP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2VLP(T, S, V, Range)
    If Range = 0 Then
        US_VLP_TS = "Error!"
    Else
        US_VLP_TS = V / 0.062428
    
End If
End Function

Function US_VHP_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2VHP(T, S, V, Range)
    If Range = 0 Then
        US_VHP_TS = "Error!"
    Else
        US_VHP_TS = V / 0.062428
    
End If
End Function

Function US_V_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2V(T, S, V, Range)
    If Range = 0 Then
        US_V_TS = "Error!"
    Else
        US_V_TS = V / 0.062428
    
End If
End Function
Function US_X_TS(ByVal T As Double, ByVal S As Double)
Rem Attribute X_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2X(T, S, X, Range)
    If Range = 0 Then
        US_X_TS = "Error!"
    Else
        US_X_TS = X
    End If
End Function
Function US_P_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute P_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2P(T, V, P, Range)
    If Range = 0 Then
        US_P_TV = "Error!"
    Else
        US_P_TV = P * 10 / 0.068948
    
End If
End Function
Function US_H_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute H_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2H(T, V, H, Range)
    If Range = 0 Then
        US_H_TV = "Error!"
    Else
        US_H_TV = H / 2.326
    End If
End Function
Function US_S_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute S_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2S(T, V, S, Range)
    If Range = 0 Then
        US_S_TV = "Error!"
    Else
        US_S_TV = S / 4.1868
    End If
End Function
Function US_X_TV(ByVal T As Double, ByVal V As Double)
Rem Attribute X_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2X(T, V, X, Range)
    If Range = 0 Then
        US_X_TV = "Error!"
    Else
        US_X_TV = X
    End If
End Function
Function US_P_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute P_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2P(T, X, P, Range)
    If Range = 0 Then
        US_P_TX = "Error!"
    Else
        US_P_TX = P * 10 / 0.068948
    
End If
End Function
Function US_H_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute H_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2H(T, X, H, Range)
    If Range = 0 Then
        US_H_TX = "Error!"
    Else
        US_H_TX = H / 2.326
    End If
End Function
Function US_S_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute S_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2S(T, X, S, Range)
    If Range = 0 Then
        US_S_TX = "Error!"
    Else
        US_S_TX = S / 4.1868
    End If
End Function
Function US_V_TX(ByVal T As Double, ByVal X As Double)
Rem Attribute V_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2V(T, X, V, Range)
    If Range = 0 Then
        US_V_TX = "Error!"
    Else
        US_V_TX = V / 0.062428
    
End If
End Function


Function US_P_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute P_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2P(H, S, P, Range)
    If Range = 0 Then
        US_P_HS = "Error!"
    Else
        US_P_HS = P * 10 / 0.068948
    
End If
End Function

Function US_T_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute T_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的温度T(F)?"
Rem Attribute T_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2T(H, S, T, Range)
    If Range = 0 Then
        US_T_HS = "Error!"
    Else
        US_T_HS = T * 1.8 + 32
    End If
End Function

Function US_V_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute V_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2V(H, S, V, Range)
    If Range = 0 Then
        US_V_HS = "Error!"
    Else
        US_V_HS = V / 0.062428
    
End If
End Function

Function US_X_HS(ByVal H As Double, ByVal S As Double)
Rem Attribute X_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2X(H, S, X, Range)
    If Range = 0 Then
        US_X_HS = "Error!"
    Else
        US_X_HS = X
    End If
End Function

Function US_P_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute P_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2P(H, V, P, Range)
    If Range = 0 Then
        US_P_HV = "Error!"
    Else
        US_P_HV = P * 10 / 0.068948
    
End If
End Function

Function US_T_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute T_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2T(H, V, T, Range)
    If Range = 0 Then
        US_T_HV = "Error!"
    Else
        US_T_HV = T * 1.8 + 32
    End If
End Function

Function US_S_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute S_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2S(H, V, S, Range)
    If Range = 0 Then
        US_S_HV = "Error!"
    Else
        US_S_HV = S / 4.1868
    End If
End Function

Function US_X_HV(ByVal H As Double, ByVal V As Double)
Rem Attribute X_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2X(H, V, X, Range)
    If Range = 0 Then
        US_X_HV = "Error!"
    Else
        US_X_HV = X
    End If
End Function

Function US_P_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2P(H, X, P, Range)
    If Range = 0 Then
        US_P_HX = "Error!"
    Else
        US_P_HX = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2PLP(H, X, P, Range)
    If Range = 0 Then
        US_PLP_HX = "Error!"
    Else
        US_PLP_HX = P * 10 / 0.068948
    
End If
End Function

Function US_PHP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2PHP(H, X, P, Range)
    If Range = 0 Then
        US_PHP_HX = "Error!"
    Else
        US_PHP_HX = P * 10 / 0.068948
    
End If
End Function


Function US_T_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2T(H, X, T, Range)
    If Range = 0 Then
        US_T_HX = "Error!"
    Else
        US_T_HX = T * 1.8 + 32
    End If
End Function

Function US_TLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2TLP(H, X, T, Range)
    If Range = 0 Then
        US_TLP_HX = "Error!"
    Else
        US_TLP_HX = T * 1.8 + 32
    End If
End Function

Function US_THP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2THP(H, X, T, Range)
    If Range = 0 Then
        US_THP_HX = "Error!"
    Else
        US_THP_HX = T * 1.8 + 32
    End If
End Function

Function US_S_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2S(H, X, S, Range)
    If Range = 0 Then
        US_S_HX = "Error!"
    Else
        US_S_HX = S / 4.1868
    End If
End Function

Function US_SLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2SLP(H, X, S, Range)
    If Range = 0 Then
        US_SLP_HX = "Error!"
    Else
        US_SLP_HX = S / 4.1868
    End If
End Function

Function US_SHP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2SHP(H, X, S, Range)
    If Range = 0 Then
        US_SHP_HX = "Error!"
    Else
        US_SHP_HX = S / 4.1868
    End If
End Function

Function US_V_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2V(H, X, V, Range)
    If Range = 0 Then
        US_V_HX = "Error!"
    Else
        US_V_HX = V / 0.062428
    
End If
End Function


Function US_VLP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2VLP(H, X, V, Range)
    If Range = 0 Then
        US_VLP_HX = "Error!"
    Else
        US_VLP_HX = V / 0.062428
    
End If
End Function


Function US_VHP_HX(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2VHP(H, X, V, Range)
    If Range = 0 Then
        US_VHP_HX = "Error!"
    Else
        US_VHP_HX = V / 0.062428
    
End If
End Function


Function US_P_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute P_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2P(S, V, P, Range)
    If Range = 0 Then
        US_P_SV = "Error!"
    Else
        US_P_SV = P * 10 / 0.068948
    
End If
End Function

Function US_T_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute T_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2T(S, V, T, Range)
    If Range = 0 Then
        US_T_SV = "Error!"
    Else
        US_T_SV = T * 1.8 + 32
    End If
End Function

Function US_H_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute H_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2H(S, V, H, Range)
    If Range = 0 Then
        US_H_SV = "Error!"
    Else
        US_H_SV = H / 2.326
    End If
End Function

Function US_X_SV(ByVal S As Double, ByVal V As Double)
Rem Attribute X_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2X(S, V, X, Range)
    If Range = 0 Then
        US_X_SV = "Error!"
    Else
        US_X_SV = X
    End If
End Function

Function US_P_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2P(S, X, P, Range)
    If Range = 0 Then
        US_P_SX = "Error!"
    Else
        US_P_SX = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PLP(S, X, P, Range)
    If Range = 0 Then
        US_PLP_SX = "Error!"
    Else
        US_PLP_SX = P * 10 / 0.068948
    
End If
End Function


Function US_PMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PMP(S, X, P, Range)
    If Range = 0 Then
        US_PMP_SX = "Error!"
    Else
        US_PMP_SX = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PHP(S, X, P, Range)
    If Range = 0 Then
        US_PHP_SX = "Error!"
    Else
        US_PHP_SX = P * 10 / 0.068948
    
End If
End Function


Function US_T_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2T(S, X, T, Range)
    If Range = 0 Then
        US_T_SX = "Error!"
    Else
        US_T_SX = T * 1.8 + 32
    End If
End Function

Function US_TLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2TLP(S, X, T, Range)
    If Range = 0 Then
        US_TLP_SX = "Error!"
    Else
        US_TLP_SX = T * 1.8 + 32
    End If
End Function

Function US_TMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2TMP(S, X, T, Range)
    If Range = 0 Then
        US_TMP_SX = "Error!"
    Else
        US_TMP_SX = T * 1.8 + 32
    End If
End Function

Function US_THP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2THP(S, X, T, Range)
    If Range = 0 Then
        US_THP_SX = "Error!"
    Else
        US_THP_SX = T * 1.8 + 32
    End If
End Function

Function US_H_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2H(S, X, H, Range)
    If Range = 0 Then
        US_H_SX = "Error!"
    Else
        US_H_SX = H / 2.326
    End If
End Function

Function US_HLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HLP(S, X, H, Range)
    If Range = 0 Then
        US_HLP_SX = "Error!"
    Else
        US_HLP_SX = H / 2.326
    End If
End Function

Function US_HMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HMP(S, X, H, Range)
    If Range = 0 Then
        US_HMP_SX = "Error!"
    Else
        US_HMP_SX = H / 2.326
    End If
End Function

Function US_HHP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HHP(S, X, H, Range)
    If Range = 0 Then
        US_HHP_SX = "Error!"
    Else
        US_HHP_SX = H / 2.326
    End If
End Function

Function US_V_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2V(S, X, V, Range)
    If Range = 0 Then
        US_V_SX = "Error!"
    Else
        US_V_SX = V / 0.062428
    
End If
End Function

Function US_VLP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VLP(S, X, V, Range)
    If Range = 0 Then
        US_VLP_SX = "Error!"
    Else
        US_VLP_SX = V / 0.062428
    
End If
End Function

Function US_VMP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VMP(S, X, V, Range)
    If Range = 0 Then
        US_VMP_SX = "Error!"
    Else
        US_VMP_SX = V / 0.062428
    
End If
End Function

Function US_VHP_SX(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VHP(S, X, V, Range)
    If Range = 0 Then
        US_VHP_SX = "Error!"
    Else
        US_VHP_SX = V / 0.062428
    
End If
End Function

Function US_P_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2P(V, X, P, Range)
    If Range = 0 Then
        US_P_VX = "Error!"
    Else
        US_P_VX = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2PLP(V, X, P, Range)
    If Range = 0 Then
        US_PLP_VX = "Error!"
    Else
        US_PLP_VX = P * 10 / 0.068948
    
End If
End Function

Function US_PHP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2PHP(V, X, P, Range)
    If Range = 0 Then
        US_PHP_VX = "Error!"
    Else
        US_PHP_VX = P * 10 / 0.068948
    
End If
End Function

Function US_T_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2T(V, X, T, Range)
    If Range = 0 Then
        US_T_VX = "Error!"
    Else
        US_T_VX = T * 1.8 + 32
    End If
End Function

Function US_TLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2TLP(V, X, T, Range)
    If Range = 0 Then
        US_TLP_VX = "Error!"
    Else
        US_TLP_VX = T * 1.8 + 32
    End If
End Function


Function US_THP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2THP(V, X, T, Range)
    If Range = 0 Then
        US_THP_VX = "Error!"
    Else
        US_THP_VX = T * 1.8 + 32
    End If
End Function


Function US_H_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2H(V, X, H, Range)
    If Range = 0 Then
        US_H_VX = "Error!"
    Else
        US_H_VX = H / 2.326
    End If
End Function

Function US_HLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2HLP(V, X, H, Range)
    If Range = 0 Then
        US_HLP_VX = "Error!"
    Else
        US_HLP_VX = H / 2.326
    End If
End Function

Function US_HHP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2HHP(V, X, H, Range)
    If Range = 0 Then
        US_HHP_VX = "Error!"
    Else
        US_HHP_VX = H / 2.326
    End If
End Function

Function US_S_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2S(V, X, S, Range)
    If Range = 0 Then
        US_S_VX = "Error!"
    Else
        US_S_VX = S / 4.1868
    End If
End Function

Function US_SLP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2SLP(V, X, S, Range)
    If Range = 0 Then
        US_SLP_VX = "Error!"
    Else
        US_SLP_VX = S / 4.1868
    End If
End Function

Function US_SHP_VX(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2SHP(V, X, S, Range)
    If Range = 0 Then
        US_SHP_VX = "Error!"
    Else
        US_SHP_VX = S / 4.1868
    End If
End Function





Rem *************************************************************************************


Function US_T_P67(ByVal P As Double)
Rem Attribute T_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和温度T(F)?"
Rem Attribute T_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2T67(P, T, Range)
    If Range = 0 Then
        US_T_P67 = "Error!"
    Else
        US_T_P67 = T * 1.8 + 32
    End If
End Function


Function US_HL_P67(ByVal P As Double)
Rem Attribute Hw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比焓Hw(Btu/lbm)?"
Rem Attribute Hw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2HL67(P, H, Range)
    If Range = 0 Then
        US_HL_P67 = "Error!"
    Else
        US_HL_P67 = H / 2.326
    End If
End Function

Function US_HG_P67(ByVal P As Double)
Rem Attribute Hs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比焓Hs(Btu/lbm)?"
Rem Attribute Hs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2HG67(P, H, Range)
    If Range = 0 Then
        US_HG_P67 = "Error!"
    Else
        US_HG_P67 = H / 2.326
    End If
End Function

Function US_SL_P67(ByVal P As Double)
Rem Attribute Sw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比熵Sw( (Btu/lbmR) )?"
Rem Attribute Sw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SL67(P, S, Range)
    If Range = 0 Then
        US_SL_P67 = "Error!"
    Else
        US_SL_P67 = S / 4.1868
    End If
End Function

Function US_SG_P67(ByVal P As Double)
Rem Attribute Ss_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比熵Ss( (Btu/lbmR) )?"
Rem Attribute Ss_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SG67(P, S, Range)
    If Range = 0 Then
        US_SG_P67 = "Error!"
    Else
        US_SG_P67 = S / 4.1868
    End If
End Function


Function US_VL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2VL67(P, V, Range)
    If Range = 0 Then
        US_VL_P67 = "Error!"
    Else
        US_VL_P67 = V / 0.062428
    
End If
End Function

Function US_VG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2VG67(P, V, Range)
    If Range = 0 Then
        US_VG_P67 = "Error!"
    Else
        US_VG_P67 = V / 0.062428
    
End If
End Function


Function US_CPL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CPL67(P, CP, Range)
    If Range = 0 Then
        US_CPL_P67 = "Error!"
    Else
        US_CPL_P67 = CP
    End If
End Function

Function US_CPG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CPG67(P, CP, Range)
    If Range = 0 Then
        US_CPG_P67 = "Error!"
    Else
        US_CPG_P67 = CP
    End If
End Function

Function US_CVL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CVL67(P, CV, Range)
    If Range = 0 Then
        US_CVL_P67 = "Error!"
    Else
        US_CVL_P67 = CV
    End If
End Function

Function US_CVG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CVG67(P, CV, Range)
    If Range = 0 Then
        US_CVG_P67 = "Error!"
    Else
        US_CVG_P67 = CV
    End If
End Function

Function US_EL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EL67(P, E, Range)
    If Range = 0 Then
        US_EL_P67 = "Error!"
    Else
        US_EL_P67 = E
    End If
End Function

Function US_EG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EG67(P, E, Range)
    If Range = 0 Then
        US_EG_P67 = "Error!"
    Else
        US_EG_P67 = E
    End If
End Function

Function US_SSPL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SSPL67(P, SSP, Range)
    If Range = 0 Then
        US_SSPL_P67 = "Error!"
    Else
        US_SSPL_P67 = SSP
    End If
End Function

Function US_SSPG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SSPG67(P, SSP, Range)
    If Range = 0 Then
        US_SSPG_P67 = "Error!"
    Else
        US_SSPG_P67 = SSP
    End If
End Function

Function US_KSL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2KSL67(P, KS, Range)
    If Range = 0 Then
        US_KSL_P67 = "Error!"
    Else
        US_KSL_P67 = KS
    End If
End Function

Function US_KSG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2KSG67(P, KS, Range)
    If Range = 0 Then
        US_KSG_P67 = "Error!"
    Else
        US_KSG_P67 = KS
    End If
End Function


Function US_ETAL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2ETAL67(P, ETA, Range)
    If Range = 0 Then
        US_ETAL_P67 = "Error!"
    Else
        US_ETAL_P67 = ETA
    End If
End Function

Function US_ETAG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2ETAG67(P, ETA, Range)
    If Range = 0 Then
        US_ETAG_P67 = "Error!"
    Else
        US_ETAG_P67 = ETA
    End If
End Function

Function US_UL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2UL67(P, U, Range)
    If Range = 0 Then
        US_UL_P67 = "Error!"
    Else
        US_UL_P67 = U
    End If
End Function

Function US_UG_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2UG67(P, U, Range)
    If Range = 0 Then
        US_UG_P67 = "Error!"
    Else
        US_UG_P67 = U
    End If
End Function

Function US_RAMDL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2RAMDL67(P, RAMD, Range)
    If Range = 0 Then
        US_RAMDL_P67 = "Error!"
    Else
        US_RAMDL_P67 = RAMD
    End If
End Function

Function US_RAMDG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2RAMDG67(P, RAMD, Range)
    If Range = 0 Then
        US_RAMDG_P67 = "Error!"
    Else
        US_RAMDG_P67 = RAMD
    End If
End Function


Function US_PRNL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2PRNL67(P, PRN, Range)
    If Range = 0 Then
        US_PRNL_P67 = "Error!"
    Else
        US_PRNL_P67 = PRN
    End If
End Function

Function US_PRNG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2PRNG67(P, PRN, Range)
    If Range = 0 Then
        US_PRNG_P67 = "Error!"
    Else
        US_PRNG_P67 = PRN
    End If
End Function


Function US_EPSL_P67(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EPSL67(P, eps, Range)
    If Range = 0 Then
        US_EPSL_P67 = "Error!"
    Else
        US_EPSL_P67 = eps
    End If
End Function

Function US_EPSG_P67(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EPSG67(P, eps, Range)
    If Range = 0 Then
        US_EPSG_P67 = "Error!"
    Else
        US_EPSG_P67 = eps
    End If
End Function

Function US_NL_P67(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2NL67(P, Lamd, n, Range)
    If Range = 0 Then
        US_NL_P67 = "Error!"
    Else
        US_NL_P67 = n
    End If
End Function

Function US_NG_P67(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2NG67(P, Lamd, n, Range)
    If Range = 0 Then
        US_NG_P67 = "Error!"
    Else
        US_NG_P67 = n
    End If
End Function

Function US_H_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute H_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2H67(P, T, H, Range)
    If Range = 0 Then
        US_H_PT67 = "Error!"
    Else
        US_H_PT67 = H / 2.326
    End If
End Function
Function US_S_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute S_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2S67(P, T, S, Range)
    If Range = 0 Then
        US_S_PT67 = "Error!"
    Else
        US_S_PT67 = S / 4.1868
    End If
End Function
Function US_V_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute V_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2V67(P, T, V, Range)
    If Range = 0 Then
        US_V_PT67 = "Error!"
    Else
        US_V_PT67 = V / 0.062428
    
End If
End Function
Function US_X_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute X_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的干度X(100%)?"
Rem Attribute X_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2X67(P, T, X, Range)
    If Range = 0 Then
        US_X_PT67 = "Error!"
    Else
        US_X_PT67 = X
    End If
End Function
Function US_ETA_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2ETA67(P, T, ETA, Range)
    If Range = 0 Then
        US_ETA_PT67 = "Error!"
    Else
        US_ETA_PT67 = ETA
    End If
End Function

Function US_U_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2U67(P, T, U, Range)
    If Range = 0 Then
        US_U_PT67 = "Error!"
    Else
        US_U_PT67 = U
    End If
End Function


Function US_RAMD_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute RAMD_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的导热系数Ramd( mW/(m.℃) )?"
Rem Attribute RAMD_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求RAMD(导热系数)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2RAMD67(P, T, RAMD, Range)
    If Range = 0 Then
        US_RAMD_PT67 = "Error!"
    Else
        US_RAMD_PT67 = RAMD
    End If
End Function

Function US_CP_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2CP67(P, T, CP, Range)
    If Range = 0 Then
        US_CP_PT67 = "Error!"
    Else
        US_CP_PT67 = CP
    End If
End Function

Function US_CV_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2CV67(P, T, CV, Range)
    If Range = 0 Then
        US_CV_PT67 = "Error!"
    Else
        US_CV_PT67 = CV
    End If
End Function

Function US_E_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2E67(P, T, E, Range)
    If Range = 0 Then
        US_E_PT67 = "Error!"
    Else
        US_E_PT67 = E
    End If
End Function
Function US_KS_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute K_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的绝热指数K(100%)?"
Rem Attribute K_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求K(绝热指数)。
    Dim K As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2KS67(P, T, K, Range)
    If Range = 0 Then
        US_KS_PT67 = "Error!"
    Else
        US_KS_PT67 = K
    End If
End Function

Function US_SSP_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute A_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的音速A (m/s)?"
Rem Attribute A_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求A(音速)。
    Dim a As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2SSP67(P, T, a, Range)
    If Range = 0 Then
        US_SSP_PT67 = "Error!"
    Else
        US_SSP_PT67 = a
    End If
End Function

Function US_PRN_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2PRN67(P, T, PRN, Range)
    If Range = 0 Then
        US_PRN_PT67 = "Error!"
    Else
        US_PRN_PT67 = PRN
    End If
End Function

Function US_EPS_PT67(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2EPS67(P, T, eps, Range)
    If Range = 0 Then
        US_EPS_PT67 = "Error!"
    Else
        US_EPS_PT67 = eps
    End If
End Function

Function US_N_PT67(ByVal P As Double, ByVal T As Double, ByVal Lamd As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2N67(P, T, Lamd, n, Range)
    If Range = 0 Then
        US_N_PT67 = "Error!"
    Else
        US_N_PT67 = n
    End If
End Function

Function US_T_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute T_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2T67(P, H, T, Range)
    If Range = 0 Then
        US_T_PH67 = "Error!"
    Else
        US_T_PH67 = T * 1.8 + 32
    End If
End Function
Function US_S_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute S_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2S67(P, H, S, Range)
    If Range = 0 Then
        US_S_PH67 = "Error!"
    Else
        US_S_PH67 = S / 4.1868
    End If
End Function
Function US_V_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute V_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2V67(P, H, V, Range)
    If Range = 0 Then
        US_V_PH67 = "Error!"
    Else
        US_V_PH67 = V / 0.062428
    
End If
End Function
Function US_X_PH67(ByVal P As Double, ByVal H As Double)
Rem Attribute X_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2X67(P, H, X, Range)
    If Range = 0 Then
        US_X_PH67 = "Error!"
    Else
        US_X_PH67 = X
    End If
End Function


Function US_T_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute T_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的温度T(F)?"
Rem Attribute T_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2T67(P, S, T, Range)
    If Range = 0 Then
        US_T_PS67 = "Error!"
    Else
        US_T_PS67 = T * 1.8 + 32
    End If
End Function
Function US_H_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute H_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2H67(P, S, H, Range)
    If Range = 0 Then
        US_H_PS67 = "Error!"
    Else
        US_H_PS67 = H / 2.326
    End If
End Function
Function US_V_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute V_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2V67(P, S, V, Range)
    If Range = 0 Then
        US_V_PS67 = "Error!"
    Else
        US_V_PS67 = V / 0.062428
    
End If
End Function
Function US_X_PS67(ByVal P As Double, ByVal S As Double)
Rem Attribute X_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2X67(P, S, X, Range)
    If Range = 0 Then
        US_X_PS67 = "Error!"
    Else
        US_X_PS67 = X
    End If
End Function


Function US_T_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute T_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2T67(P, V, T, Range)
    If Range = 0 Then
        US_T_PV67 = "Error!"
    Else
        US_T_PV67 = T * 1.8 + 32
    End If
End Function
Function US_H_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute H_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2H67(P, V, H, Range)
    If Range = 0 Then
        US_H_PV67 = "Error!"
    Else
        US_H_PV67 = H / 2.326
    End If
End Function
Function US_S_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute S_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2S67(P, V, S, Range)
    If Range = 0 Then
        US_S_PV67 = "Error!"
    Else
        US_S_PV67 = S / 4.1868
    End If
End Function
Function US_X_PV67(ByVal P As Double, ByVal V As Double)
Rem Attribute X_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),求对应的干度X(100%)?"
Rem Attribute X_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2X67(P, V, X, Range)
    If Range = 0 Then
        US_X_PV67 = "Error!"
    Else
        US_X_PV67 = X
    End If
End Function
Function US_T_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute T_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2T67(P, X, T, Range)
    If Range = 0 Then
        US_T_PX67 = "Error!"
    Else
        US_T_PX67 = T * 1.8 + 32
    End If
End Function
Function US_H_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute H_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2H67(P, X, H, Range)
    If Range = 0 Then
        US_H_PX67 = "Error!"
    Else
        US_H_PX67 = H / 2.326
    End If
End Function
Function US_S_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute S_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2S67(P, X, S, Range)
    If Range = 0 Then
        US_S_PX67 = "Error!"
    Else
        US_S_PX67 = S / 4.1868
    End If
End Function
Function US_V_PX67(ByVal P As Double, ByVal X As Double)
Rem Attribute V_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2V67(P, X, V, Range)
    If Range = 0 Then
        US_V_PX67 = "Error!"
    Else
        US_V_PX67 = V / 0.062428
    
End If
End Function


Function US_P_T67(ByVal T As Double)
Rem Attribute P_T.VB_Description = "已知温度T(F),\r\n求对应的饱和压力P(Psi)?"
Rem Attribute P_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2P67(T, P, Range)
    If Range = 0 Then
        US_P_T67 = "Error!"
    Else
        US_P_T67 = P * 10 / 0.068948
    
End If
End Function
Function US_HL_T67(ByVal T As Double)
Rem Attribute Hw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比焓Hw(Btu/lbm)?"
Rem Attribute Hw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2HL67(T, H, Range)
    If Range = 0 Then
        US_HL_T67 = "Error!"
    Else
        US_HL_T67 = H / 2.326
    End If
End Function
Function US_HG_T67(ByVal T As Double)
Rem Attribute Hs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比焓Hs(Btu/lbm)?"
Rem Attribute Hs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HS(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2HG67(T, H, Range)
    If Range = 0 Then
        US_HG_T67 = "Error!"
    Else
        US_HG_T67 = H / 2.326
    End If
End Function
Function US_SG_T67(ByVal T As Double)
Rem Attribute Ss_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比熵Ss( (Btu/lbmR) )?"
Rem Attribute Ss_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SS(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SG67(T, S, Range)
    If Range = 0 Then
        US_SG_T67 = "Error!"
    Else
        US_SG_T67 = S / 4.1868
    End If
End Function
Function US_SL_T67(ByVal T As Double)
Rem Attribute Sw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比熵Sw( (Btu/lbmR) )?"
Rem Attribute Sw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SL67(T, S, Range)
    If Range = 0 Then
        US_SL_T67 = "Error!"
    Else
        US_SL_T67 = S / 4.1868
    End If
End Function
Function US_VL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2VL67(T, V, Range)
    If Range = 0 Then
        US_VL_T67 = "Error!"
    Else
        US_VL_T67 = V / 0.062428
    
End If
End Function
Function US_VG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2VG67(T, V, Range)
    If Range = 0 Then
        US_VG_T67 = "Error!"
    Else
        US_VG_T67 = V / 0.062428
    
End If
End Function


Function US_CPL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CPL67(T, CP, Range)
    If Range = 0 Then
        US_CPL_T67 = "Error!"
    Else
        US_CPL_T67 = CP
    End If
End Function
Function US_CPG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CPG67(T, CP, Range)
    If Range = 0 Then
        US_CPG_T67 = "Error!"
    Else
        US_CPG_T67 = CP
    End If
End Function


Function US_CVL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CVL67(T, CV, Range)
    If Range = 0 Then
        US_CVL_T67 = "Error!"
    Else
        US_CVL_T67 = CV
    End If
End Function
Function US_CVG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CVG67(T, CV, Range)
    If Range = 0 Then
        US_CVG_T67 = "Error!"
    Else
        US_CVG_T67 = CV
    End If
End Function

Function US_EL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EL67(T, E, Range)
    If Range = 0 Then
        US_EL_T67 = "Error!"
    Else
        US_EL_T67 = E
    End If
End Function
Function US_EG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EG67(T, E, Range)
    If Range = 0 Then
        US_EG_T67 = "Error!"
    Else
        US_EG_T67 = E
    End If
End Function

Function US_SSPL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SSPL67(T, SSP, Range)
    If Range = 0 Then
        US_SSPL_T67 = "Error!"
    Else
        US_SSPL_T67 = SSP
    End If
End Function
Function US_SSPG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SSPG67(T, SSP, Range)
    If Range = 0 Then
        US_SSPG_T67 = "Error!"
    Else
        US_SSPG_T67 = SSP
    End If
End Function



Function US_KSL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2KSL67(T, KS, Range)
    If Range = 0 Then
        US_KSL_T67 = "Error!"
    Else
        US_KSL_T67 = KS
    End If
End Function
Function US_KSG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2KSG67(T, KS, Range)
    If Range = 0 Then
        US_KSG_T67 = "Error!"
    Else
        US_KSG_T67 = KS
    End If
End Function


Function US_ETAL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2ETAL67(T, ETA, Range)
    If Range = 0 Then
        US_ETAL_T67 = "Error!"
    Else
        US_ETAL_T67 = ETA
    End If
End Function
Function US_ETAG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2ETAG67(T, ETA, Range)
    If Range = 0 Then
        US_ETAG_T67 = "Error!"
    Else
        US_ETAG_T67 = ETA
    End If
End Function

Function US_UL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2UL67(T, U, Range)
    If Range = 0 Then
        US_UL_T67 = "Error!"
    Else
        US_UL_T67 = U
    End If
End Function

Function US_UG_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2UG67(T, U, Range)
    If Range = 0 Then
        US_UG_T67 = "Error!"
    Else
        US_UG_T67 = U
    End If
End Function

Function US_RAMDL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2RAMDL67(T, RAMD, Range)
    If Range = 0 Then
        US_RAMDL_T67 = "Error!"
    Else
        US_RAMDL_T67 = RAMD
    End If
End Function
Function US_RAMDG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2RAMDG67(T, RAMD, Range)
    If Range = 0 Then
        US_RAMDG_T67 = "Error!"
    Else
        US_RAMDG_T67 = RAMD
    End If
End Function




Function US_PRNL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2PRNL67(T, PRN, Range)
    If Range = 0 Then
        US_PRNL_T67 = "Error!"
    Else
        US_PRNL_T67 = PRN
    End If
End Function
Function US_PRNG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2PRNG67(T, PRN, Range)
    If Range = 0 Then
        US_PRNG_T67 = "Error!"
    Else
        US_PRNG_T67 = PRN
    End If
End Function

Function US_EPSL_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EPSL67(T, eps, Range)
    If Range = 0 Then
        US_EPSL_T67 = "Error!"
    Else
        US_EPSL_T67 = eps
    End If
End Function
Function US_EPSG_T67(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EPSG67(T, eps, Range)
    If Range = 0 Then
        US_EPSG_T67 = "Error!"
    Else
        US_EPSG_T67 = eps
    End If
End Function

Function US_NL_T67(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2NL67(T, Lamd, n, Range)
    If Range = 0 Then
        US_NL_T67 = "Error!"
    Else
        US_NL_T67 = n
    End If
End Function

Function US_NG_T67(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2NG67(T, Lamd, n, Range)
    If Range = 0 Then
        US_NG_T67 = "Error!"
    Else
        US_NG_T67 = n
    End If
End Function

Function US_SurfT_T67(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SurfT As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SURFT67(T, SurfT, Range)
    If Range = 0 Then
        US_SurfT_T67 = "Error!"
    Else
        US_SurfT_T67 = SurfT
    End If
End Function

Function US_P_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2P67(T, H, P, Range)
    If Range = 0 Then
        US_P_TH67 = "Error!"
    Else
        US_P_TH67 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2PLP67(T, H, P, Range)
    If Range = 0 Then
        US_PLP_TH67 = "Error!"
    Else
        US_PLP_TH67 = P * 10 / 0.068948
    
End If
End Function



Function US_PHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2PHP67(T, H, P, Range)
    If Range = 0 Then
        US_PHP_TH67 = "Error!"
    Else
        US_PHP_TH67 = P * 10 / 0.068948
    
End If
End Function

Function US_S_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2S67(T, H, S, Range)
    If Range = 0 Then
        US_S_TH67 = "Error!"
    Else
        US_S_TH67 = S / 4.1868
    End If
End Function

Function US_SLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2SLP67(T, H, S, Range)
    If Range = 0 Then
        US_SLP_TH67 = "Error!"
    Else
        US_SLP_TH67 = S / 4.1868
    End If
End Function

Function US_SHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2SHP67(T, H, S, Range)
    If Range = 0 Then
        US_SHP_TH67 = "Error!"
    Else
        US_SHP_TH67 = S / 4.1868
    End If
End Function


Function US_V_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2V67(T, H, V, Range)
    If Range = 0 Then
        US_V_TH67 = "Error!"
    Else
        US_V_TH67 = V / 0.062428
    
End If
End Function


Function US_VLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2VLP67(T, H, V, Range)
    If Range = 0 Then
        US_VLP_TH67 = "Error!"
    Else
        US_VLP_TH67 = V / 0.062428
    
End If
End Function


Function US_VHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2VHP67(T, H, V, Range)
    If Range = 0 Then
        US_VHP_TH67 = "Error!"
    Else
        US_VHP_TH67 = V / 0.062428
    
End If
End Function

Function US_XLP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2XLP67(T, H, X, Range)
    If Range = 0 Then
        US_XLP_TH67 = "Error!"
    Else
        US_XLP_TH67 = X
    End If
End Function
Function US_XHP_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2XHP67(T, H, X, Range)
    If Range = 0 Then
        US_XHP_TH67 = "Error!"
    Else
        US_XHP_TH67 = X
    End If
End Function
Function US_X_TH67(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2X67(T, H, X, Range)
    If Range = 0 Then
        US_X_TH67 = "Error!"
    Else
        US_X_TH67 = X
    End If
End Function


Function US_PLP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2PLP67(T, S, P, Range)
    If Range = 0 Then
        US_PLP_TS67 = "Error!"
    Else
        US_PLP_TS67 = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2PHP67(T, S, P, Range)
    If Range = 0 Then
        US_PHP_TS67 = "Error!"
    Else
        US_PHP_TS67 = P * 10 / 0.068948
    
End If
End Function
Function US_P_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2P67(T, S, P, Range)
    If Range = 0 Then
        US_P_TS67 = "Error!"
    Else
        US_P_TS67 = P * 10 / 0.068948
    
End If
End Function
Function US_HLP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2HLP67(T, S, H, Range)
    If Range = 0 Then
        US_HLP_TS67 = "Error!"
    Else
        US_HLP_TS67 = H / 2.326
    End If
End Function


Function US_HHP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2HHP67(T, S, H, Range)
    If Range = 0 Then
        US_HHP_TS67 = "Error!"
    Else
        US_HHP_TS67 = H / 2.326
    End If
End Function
Function US_H_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2H67(T, S, H, Range)
    If Range = 0 Then
        US_H_TS67 = "Error!"
    Else
        US_H_TS67 = H / 2.326
    End If
End Function

Function US_VLP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2VLP67(T, S, V, Range)
    If Range = 0 Then
        US_VLP_TS67 = "Error!"
    Else
        US_VLP_TS67 = V / 0.062428
    
End If
End Function

Function US_VHP_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2VHP67(T, S, V, Range)
    If Range = 0 Then
        US_VHP_TS67 = "Error!"
    Else
        US_VHP_TS67 = V / 0.062428
    
End If
End Function

Function US_V_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2V67(T, S, V, Range)
    If Range = 0 Then
        US_V_TS67 = "Error!"
    Else
        US_V_TS67 = V / 0.062428
    
End If
End Function
Function US_X_TS67(ByVal T As Double, ByVal S As Double)
Rem Attribute X_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2X67(T, S, X, Range)
    If Range = 0 Then
        US_X_TS67 = "Error!"
    Else
        US_X_TS67 = X
    End If
End Function
Function US_P_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute P_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2P67(T, V, P, Range)
    If Range = 0 Then
        US_P_TV67 = "Error!"
    Else
        US_P_TV67 = P * 10 / 0.068948
    
End If
End Function
Function US_H_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute H_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2H67(T, V, H, Range)
    If Range = 0 Then
        US_H_TV67 = "Error!"
    Else
        US_H_TV67 = H / 2.326
    End If
End Function
Function US_S_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute S_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2S67(T, V, S, Range)
    If Range = 0 Then
        US_S_TV67 = "Error!"
    Else
        US_S_TV67 = S / 4.1868
    End If
End Function
Function US_X_TV67(ByVal T As Double, ByVal V As Double)
Rem Attribute X_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2X67(T, V, X, Range)
    If Range = 0 Then
        US_X_TV67 = "Error!"
    Else
        US_X_TV67 = X
    End If
End Function
Function US_P_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute P_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2P67(T, X, P, Range)
    If Range = 0 Then
        US_P_TX67 = "Error!"
    Else
        US_P_TX67 = P * 10 / 0.068948
    
End If
End Function
Function US_H_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute H_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2H67(T, X, H, Range)
    If Range = 0 Then
        US_H_TX67 = "Error!"
    Else
        US_H_TX67 = H / 2.326
    End If
End Function
Function US_S_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute S_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2S67(T, X, S, Range)
    If Range = 0 Then
        US_S_TX67 = "Error!"
    Else
        US_S_TX67 = S / 4.1868
    End If
End Function
Function US_V_TX67(ByVal T As Double, ByVal X As Double)
Rem Attribute V_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2V67(T, X, V, Range)
    If Range = 0 Then
        US_V_TX67 = "Error!"
    Else
        US_V_TX67 = V / 0.062428
    
End If
End Function


Function US_P_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute P_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2P67(H, S, P, Range)
    If Range = 0 Then
        US_P_HS67 = "Error!"
    Else
        US_P_HS67 = P * 10 / 0.068948
    
End If
End Function

Function US_T_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute T_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的温度T(F)?"
Rem Attribute T_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2T67(H, S, T, Range)
    If Range = 0 Then
        US_T_HS67 = "Error!"
    Else
        US_T_HS67 = T * 1.8 + 32
    End If
End Function

Function US_V_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute V_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2V67(H, S, V, Range)
    If Range = 0 Then
        US_V_HS67 = "Error!"
    Else
        US_V_HS67 = V / 0.062428
    
End If
End Function

Function US_X_HS67(ByVal H As Double, ByVal S As Double)
Rem Attribute X_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2X67(H, S, X, Range)
    If Range = 0 Then
        US_X_HS67 = "Error!"
    Else
        US_X_HS67 = X
    End If
End Function

Function US_P_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute P_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2P67(H, V, P, Range)
    If Range = 0 Then
        US_P_HV67 = "Error!"
    Else
        US_P_HV67 = P * 10 / 0.068948
    
End If
End Function

Function US_T_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute T_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2T67(H, V, T, Range)
    If Range = 0 Then
        US_T_HV67 = "Error!"
    Else
        US_T_HV67 = T * 1.8 + 32
    End If
End Function

Function US_S_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute S_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2S67(H, V, S, Range)
    If Range = 0 Then
        US_S_HV67 = "Error!"
    Else
        US_S_HV67 = S / 4.1868
    End If
End Function

Function US_X_HV67(ByVal H As Double, ByVal V As Double)
Rem Attribute X_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2X67(H, V, X, Range)
    If Range = 0 Then
        US_X_HV67 = "Error!"
    Else
        US_X_HV67 = X
    End If
End Function

Function US_P_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2P67(H, X, P, Range)
    If Range = 0 Then
        US_P_HX67 = "Error!"
    Else
        US_P_HX67 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2PLP67(H, X, P, Range)
    If Range = 0 Then
        US_PLP_HX67 = "Error!"
    Else
        US_PLP_HX67 = P * 10 / 0.068948
    
End If
End Function

Function US_PHP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2PHP67(H, X, P, Range)
    If Range = 0 Then
        US_PHP_HX67 = "Error!"
    Else
        US_PHP_HX67 = P * 10 / 0.068948
    
End If
End Function


Function US_T_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2T67(H, X, T, Range)
    If Range = 0 Then
        US_T_HX67 = "Error!"
    Else
        US_T_HX67 = T * 1.8 + 32
    End If
End Function

Function US_TLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2TLP67(H, X, T, Range)
    If Range = 0 Then
        US_TLP_HX67 = "Error!"
    Else
        US_TLP_HX67 = T * 1.8 + 32
    End If
End Function

Function US_THP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2THP67(H, X, T, Range)
    If Range = 0 Then
        US_THP_HX67 = "Error!"
    Else
        US_THP_HX67 = T * 1.8 + 32
    End If
End Function

Function US_S_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2S67(H, X, S, Range)
    If Range = 0 Then
        US_S_HX67 = "Error!"
    Else
        US_S_HX67 = S / 4.1868
    End If
End Function

Function US_SLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2SLP67(H, X, S, Range)
    If Range = 0 Then
        US_SLP_HX67 = "Error!"
    Else
        US_SLP_HX67 = S / 4.1868
    End If
End Function

Function US_SHP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2SHP67(H, X, S, Range)
    If Range = 0 Then
        US_SHP_HX67 = "Error!"
    Else
        US_SHP_HX67 = S / 4.1868
    End If
End Function

Function US_V_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2V67(H, X, V, Range)
    If Range = 0 Then
        US_V_HX67 = "Error!"
    Else
        US_V_HX67 = V / 0.062428
    
End If
End Function


Function US_VLP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2VLP67(H, X, V, Range)
    If Range = 0 Then
        US_VLP_HX67 = "Error!"
    Else
        US_VLP_HX67 = V / 0.062428
    
End If
End Function


Function US_VHP_HX67(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2VHP67(H, X, V, Range)
    If Range = 0 Then
        US_VHP_HX67 = "Error!"
    Else
        US_VHP_HX67 = V / 0.062428
    
End If
End Function


Function US_P_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute P_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2P67(S, V, P, Range)
    If Range = 0 Then
        US_P_SV67 = "Error!"
    Else
        US_P_SV67 = P * 10 / 0.068948
    
End If
End Function

Function US_T_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute T_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2T67(S, V, T, Range)
    If Range = 0 Then
        US_T_SV67 = "Error!"
    Else
        US_T_SV67 = T * 1.8 + 32
    End If
End Function

Function US_H_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute H_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2H67(S, V, H, Range)
    If Range = 0 Then
        US_H_SV67 = "Error!"
    Else
        US_H_SV67 = H / 2.326
    End If
End Function

Function US_X_SV67(ByVal S As Double, ByVal V As Double)
Rem Attribute X_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2X67(S, V, X, Range)
    If Range = 0 Then
        US_X_SV67 = "Error!"
    Else
        US_X_SV67 = X
    End If
End Function

Function US_P_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2P67(S, X, P, Range)
    If Range = 0 Then
        US_P_SX67 = "Error!"
    Else
        US_P_SX67 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PLP67(S, X, P, Range)
    If Range = 0 Then
        US_PLP_SX67 = "Error!"
    Else
        US_PLP_SX67 = P * 10 / 0.068948
    
End If
End Function


Function US_PMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PMP67(S, X, P, Range)
    If Range = 0 Then
        US_PMP_SX67 = "Error!"
    Else
        US_PMP_SX67 = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PHP67(S, X, P, Range)
    If Range = 0 Then
        US_PHP_SX67 = "Error!"
    Else
        US_PHP_SX67 = P * 10 / 0.068948
    
End If
End Function


Function US_T_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2T67(S, X, T, Range)
    If Range = 0 Then
        US_T_SX67 = "Error!"
    Else
        US_T_SX67 = T * 1.8 + 32
    End If
End Function

Function US_TLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2TLP67(S, X, T, Range)
    If Range = 0 Then
        US_TLP_SX67 = "Error!"
    Else
        US_TLP_SX67 = T * 1.8 + 32
    End If
End Function

Function US_TMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2TMP67(S, X, T, Range)
    If Range = 0 Then
        US_TMP_SX67 = "Error!"
    Else
        US_TMP_SX67 = T * 1.8 + 32
    End If
End Function

Function US_THP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2THP67(S, X, T, Range)
    If Range = 0 Then
        US_THP_SX67 = "Error!"
    Else
        US_THP_SX67 = T * 1.8 + 32
    End If
End Function

Function US_H_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2H67(S, X, H, Range)
    If Range = 0 Then
        US_H_SX67 = "Error!"
    Else
        US_H_SX67 = H / 2.326
    End If
End Function

Function US_HLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HLP67(S, X, H, Range)
    If Range = 0 Then
        US_HLP_SX67 = "Error!"
    Else
        US_HLP_SX67 = H / 2.326
    End If
End Function

Function US_HMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HMP67(S, X, H, Range)
    If Range = 0 Then
        US_HMP_SX67 = "Error!"
    Else
        US_HMP_SX67 = H / 2.326
    End If
End Function

Function US_HHP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HHP67(S, X, H, Range)
    If Range = 0 Then
        US_HHP_SX67 = "Error!"
    Else
        US_HHP_SX67 = H / 2.326
    End If
End Function

Function US_V_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2V67(S, X, V, Range)
    If Range = 0 Then
        US_V_SX67 = "Error!"
    Else
        US_V_SX67 = V / 0.062428
    
End If
End Function

Function US_VLP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VLP67(S, X, V, Range)
    If Range = 0 Then
        US_VLP_SX67 = "Error!"
    Else
        US_VLP_SX67 = V / 0.062428
    
End If
End Function

Function US_VMP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VMP67(S, X, V, Range)
    If Range = 0 Then
        US_VMP_SX67 = "Error!"
    Else
        US_VMP_SX67 = V / 0.062428
    
End If
End Function

Function US_VHP_SX67(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VHP67(S, X, V, Range)
    If Range = 0 Then
        US_VHP_SX67 = "Error!"
    Else
        US_VHP_SX67 = V / 0.062428
    
End If
End Function

Function US_P_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2P67(V, X, P, Range)
    If Range = 0 Then
        US_P_VX67 = "Error!"
    Else
        US_P_VX67 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2PLP67(V, X, P, Range)
    If Range = 0 Then
        US_PLP_VX67 = "Error!"
    Else
        US_PLP_VX67 = P * 10 / 0.068948
    
End If
End Function

Function US_PHP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2PHP67(V, X, P, Range)
    If Range = 0 Then
        US_PHP_VX67 = "Error!"
    Else
        US_PHP_VX67 = P * 10 / 0.068948
    
End If
End Function

Function US_T_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2T67(V, X, T, Range)
    If Range = 0 Then
        US_T_VX67 = "Error!"
    Else
        US_T_VX67 = T * 1.8 + 32
    End If
End Function

Function US_TLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2TLP67(V, X, T, Range)
    If Range = 0 Then
        US_TLP_VX67 = "Error!"
    Else
        US_TLP_VX67 = T * 1.8 + 32
    End If
End Function


Function US_THP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2THP67(V, X, T, Range)
    If Range = 0 Then
        US_THP_VX67 = "Error!"
    Else
        US_THP_VX67 = T * 1.8 + 32
    End If
End Function


Function US_H_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2H67(V, X, H, Range)
    If Range = 0 Then
        US_H_VX67 = "Error!"
    Else
        US_H_VX67 = H / 2.326
    End If
End Function

Function US_HLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2HLP67(V, X, H, Range)
    If Range = 0 Then
        US_HLP_VX67 = "Error!"
    Else
        US_HLP_VX67 = H / 2.326
    End If
End Function

Function US_HHP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2HHP67(V, X, H, Range)
    If Range = 0 Then
        US_HHP_VX67 = "Error!"
    Else
        US_HHP_VX67 = H / 2.326
    End If
End Function

Function US_S_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2S67(V, X, S, Range)
    If Range = 0 Then
        US_S_VX67 = "Error!"
    Else
        US_S_VX67 = S / 4.1868
    End If
End Function

Function US_SLP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2SLP67(V, X, S, Range)
    If Range = 0 Then
        US_SLP_VX67 = "Error!"
    Else
        US_SLP_VX67 = S / 4.1868
    End If
End Function

Function US_SHP_VX67(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2SHP67(V, X, S, Range)
    If Range = 0 Then
        US_SHP_VX67 = "Error!"
    Else
        US_SHP_VX67 = S / 4.1868
    End If
End Function



Rem *************************************************************************************

Function US_T_P97(ByVal P As Double)
Rem Attribute T_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和温度T(F)?"
Rem Attribute T_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2T97(P, T, Range)
    If Range = 0 Then
        US_T_P97 = "Error!"
    Else
        US_T_P97 = T * 1.8 + 32
    End If
End Function


Function US_HL_P97(ByVal P As Double)
Rem Attribute Hw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比焓Hw(Btu/lbm)?"
Rem Attribute Hw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2HL97(P, H, Range)
    If Range = 0 Then
        US_HL_P97 = "Error!"
    Else
        US_HL_P97 = H / 2.326
    End If
End Function

Function US_HG_P97(ByVal P As Double)
Rem Attribute Hs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比焓Hs(Btu/lbm)?"
Rem Attribute Hs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求HW(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2HG97(P, H, Range)
    If Range = 0 Then
        US_HG_P97 = "Error!"
    Else
        US_HG_P97 = H / 2.326
    End If
End Function

Function US_SL_P97(ByVal P As Double)
Rem Attribute Sw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比熵Sw( (Btu/lbmR) )?"
Rem Attribute Sw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SL97(P, S, Range)
    If Range = 0 Then
        US_SL_P97 = "Error!"
    Else
        US_SL_P97 = S / 4.1868
    End If
End Function

Function US_SG_P97(ByVal P As Double)
Rem Attribute Ss_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比熵Ss( (Btu/lbmR) )?"
Rem Attribute Ss_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求SW(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SG97(P, S, Range)
    If Range = 0 Then
        US_SG_P97 = "Error!"
    Else
        US_SG_P97 = S / 4.1868
    End If
End Function


Function US_VL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2VL97(P, V, Range)
    If Range = 0 Then
        US_VL_P97 = "Error!"
    Else
        US_VL_P97 = V / 0.062428
    
End If
End Function

Function US_VG_P97(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2VG97(P, V, Range)
    If Range = 0 Then
        US_VG_P97 = "Error!"
    Else
        US_VG_P97 = V / 0.062428
    
End If
End Function


Function US_CpL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CPL97(P, CP, Range)
    If Range = 0 Then
        US_CpL_P97 = "Error!"
    Else
        US_CpL_P97 = CP
    End If
End Function

Function US_CpG_P97(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CPG97(P, CP, Range)
    If Range = 0 Then
        US_CpG_P97 = "Error!"
    Else
        US_CpG_P97 = CP
    End If
End Function

Function US_CvL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CVL97(P, CV, Range)
    If Range = 0 Then
        US_CvL_P97 = "Error!"
    Else
        US_CvL_P97 = CV
    End If
End Function

Function US_CvG_P97(ByVal P As Double)
Rem Attribute Vs_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和蒸汽比容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2CVG97(P, CV, Range)
    If Range = 0 Then
        US_CvG_P97 = "Error!"
    Else
        US_CvG_P97 = CV
    End If
End Function


Function US_EL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EL97(P, E, Range)
    If Range = 0 Then
        US_EL_P97 = "Error!"
    Else
        US_EL_P97 = E
    End If
End Function


Function US_EG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EG97(P, E, Range)
    If Range = 0 Then
        US_EG_P97 = "Error!"
    Else
        US_EG_P97 = E
    End If
End Function


Function US_SSpL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SSPL97(P, SSP, Range)
    If Range = 0 Then
        US_SSpL_P97 = "Error!"
    Else
        US_SSpL_P97 = SSP
    End If
End Function

Function US_SSpG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2SSPG97(P, SSP, Range)
    If Range = 0 Then
        US_SSpG_P97 = "Error!"
    Else
        US_SSpG_P97 = SSP
    End If
End Function

Function US_KsL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2KSL97(P, KS, Range)
    If Range = 0 Then
        US_KsL_P97 = "Error!"
    Else
        US_KsL_P97 = KS
    End If
End Function

Function US_KsG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2KSG97(P, KS, Range)
    If Range = 0 Then
        US_KsG_P97 = "Error!"
    Else
        US_KsG_P97 = KS
    End If
End Function

Function US_EtaL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2ETAL97(P, ETA, Range)
    If Range = 0 Then
        US_EtaL_P97 = "Error!"
    Else
        US_EtaL_P97 = ETA
    End If
End Function


Function US_EtaG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2ETAG97(P, ETA, Range)
    If Range = 0 Then
        US_EtaG_P97 = "Error!"
    Else
        US_EtaG_P97 = ETA
    End If
End Function

Function US_UL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2UL97(P, U, Range)
    If Range = 0 Then
        US_UL_P97 = "Error!"
    Else
        US_UL_P97 = U
    End If
End Function

Function US_UG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2UG97(P, U, Range)
    If Range = 0 Then
        US_UG_P97 = "Error!"
    Else
        US_UG_P97 = U
    End If
End Function

Function US_RamdL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2RAMDL97(P, RAMD, Range)
    If Range = 0 Then
        US_RamdL_P97 = "Error!"
    Else
        US_RamdL_P97 = RAMD
    End If
End Function


Function US_RamdG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2RAMDG97(P, RAMD, Range)
    If Range = 0 Then
        US_RamdG_P97 = "Error!"
    Else
        US_RamdG_P97 = RAMD
    End If
End Function

Function US_EpsL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EPSL97(P, eps, Range)
    If Range = 0 Then
        US_EpsL_P97 = "Error!"
    Else
        US_EpsL_P97 = eps
    End If
End Function

Function US_EpsG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2EPSG97(P, eps, Range)
    If Range = 0 Then
        US_EpsG_P97 = "Error!"
    Else
        US_EpsG_P97 = eps
    End If
End Function

Function US_PrnL_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2PRNL97(P, PRN, Range)
    If Range = 0 Then
        US_PrnL_P97 = "Error!"
    Else
        US_PrnL_P97 = PRN
    End If
End Function

Function US_PrnG_P97(ByVal P As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2PRNG97(P, PRN, Range)
    If Range = 0 Then
        US_PrnG_P97 = "Error!"
    Else
        US_PrnG_P97 = PRN
    End If
End Function

Function US_NL_P97(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2NL97(P, Lamd, n, Range)
    If Range = 0 Then
        US_NL_P97 = "Error!"
    Else
        US_NL_P97 = n
    End If
End Function

Function US_NG_P97(ByVal P As Double, ByVal Lamd As Double)
Rem Attribute Vw_P.VB_Description = "已知压力P(Psi),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_P.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double
    Rem P = 压力P
    P = 0.068948 * P / 10
    Call P2NG97(P, Lamd, n, Range)
    If Range = 0 Then
        US_NG_P97 = "Error!"
    Else
        US_NG_P97 = n
    End If
End Function

Function US_H_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute H_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2H97(P, T, H, Range)
    If Range = 0 Then
        US_H_PT97 = "Error!"
    Else
        US_H_PT97 = H / 2.326
    End If
End Function
Function US_S_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute S_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2S97(P, T, S, Range)
    If Range = 0 Then
        US_S_PT97 = "Error!"
    Else
        US_S_PT97 = S / 4.1868
    End If
End Function
Function US_V_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute V_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2V97(P, T, V, Range)
    If Range = 0 Then
        US_V_PT97 = "Error!"
    Else
        US_V_PT97 = V / 0.062428
    
End If
End Function
Function US_X_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute X_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的干度X(100%)?"
Rem Attribute X_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2X97(P, T, X, Range)
    If Range = 0 Then
        US_X_PT97 = "Error!"
    Else
        US_X_PT97 = X
    End If
End Function


Function US_Cp_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2CP97(P, T, CP, Range)
    If Range = 0 Then
        US_Cp_PT97 = "Error!"
    Else
        US_Cp_PT97 = CP
    End If
End Function


Function US_Cv_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim CV As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2CV97(P, T, CV, Range)
    If Range = 0 Then
        US_Cv_PT97 = "Error!"
    Else
        US_Cv_PT97 = CV
    End If
End Function

Function US_E_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim E As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2E97(P, T, E, Range)
    If Range = 0 Then
        US_E_PT97 = "Error!"
    Else
        US_E_PT97 = E
    End If
End Function


Function US_SSp_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim SSP As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2SSP97(P, T, SSP, Range)
    If Range = 0 Then
        US_SSp_PT97 = "Error!"
    Else
        US_SSp_PT97 = SSP
    End If
End Function


Function US_Ks_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute CP_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的定压质量比热容CP( (Btu/lbmR) )?"
Rem Attribute CP_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求CP(定压质量比热容)。
    Dim KS As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2KS97(P, T, KS, Range)
    If Range = 0 Then
        US_Ks_PT97 = "Error!"
    Else
        US_Ks_PT97 = KS
    End If
End Function


Function US_Eta_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim ETA As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2ETA97(P, T, ETA, Range)
    If Range = 0 Then
        US_Eta_PT97 = "Error!"
    Else
        US_Eta_PT97 = ETA
    End If
End Function

Function US_U_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute ETA_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的动力粘度Eta(10^-6 Pa.s)?"
Rem Attribute ETA_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求ETA(动力粘度)。
    Dim U As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2U97(P, T, U, Range)
    If Range = 0 Then
        US_U_PT97 = "Error!"
    Else
        US_U_PT97 = U
    End If
End Function


Function US_Ramd_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute RAMD_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的导热系数Ramd( mW/(m.℃) )?"
Rem Attribute RAMD_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求RAMD(导热系数)。
    Dim RAMD As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2RAMD97(P, T, RAMD, Range)
    If Range = 0 Then
        US_Ramd_PT97 = "Error!"
    Else
        US_Ramd_PT97 = RAMD
    End If
End Function


Function US_PRN_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim PRN As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2PRN97(P, T, PRN, Range)
    If Range = 0 Then
        US_PRN_PT97 = "Error!"
    Else
        US_PRN_PT97 = PRN
    End If
End Function

Function US_Eps_PT97(ByVal P As Double, ByVal T As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim eps As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2EPS97(P, T, eps, Range)
    If Range = 0 Then
        US_Eps_PT97 = "Error!"
    Else
        US_Eps_PT97 = eps
    End If
End Function

Function US_N_PT97(ByVal P As Double, ByVal T As Double, ByVal Lamd As Double)
Rem Attribute PRN_PT.VB_Description = "已知压力P(Psi)和温度T(F),\r\n求对应的普朗特常数PRN(100%)?"
Rem Attribute PRN_PT.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和T,求PRN(普朗特常数)。
    Dim n As Double, Range As Integer
    Rem Dim P As Double, T As Double
    Rem P = 压力P
    Rem T = 温度T
    P = 0.068948 * P / 10
    T = (T - 32) / 1.8
    Call PT2N97(P, T, Lamd, n, Range)
    If Range = 0 Then
        US_N_PT97 = "Error!"
    Else
        US_N_PT97 = n
    End If
End Function

Function US_T_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute T_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2T97(P, H, T, Range)
    If Range = 0 Then
        US_T_PH97 = "Error!"
    Else
        US_T_PH97 = T * 1.8 + 32
    End If
End Function
Function US_S_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute S_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2S97(P, H, S, Range)
    If Range = 0 Then
        US_S_PH97 = "Error!"
    Else
        US_S_PH97 = S / 4.1868
    End If
End Function
Function US_V_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute V_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2V97(P, H, V, Range)
    If Range = 0 Then
        US_V_PH97 = "Error!"
    Else
        US_V_PH97 = V / 0.062428
    
End If
End Function
Function US_X_PH97(ByVal P As Double, ByVal H As Double)
Rem Attribute X_PH.VB_Description = "已知压力P(Psi)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_PH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, H As Double
    Rem P = 压力P
    Rem H = 比焓H
    P = 0.068948 * P / 10
    H = 2.326 * H
   Call PH2X97(P, H, X, Range)
    If Range = 0 Then
        US_X_PH97 = "Error!"
    Else
        US_X_PH97 = X
    End If
End Function


Function US_T_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute T_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的温度T(F)?"
Rem Attribute T_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2T97(P, S, T, Range)
    If Range = 0 Then
        US_T_PS97 = "Error!"
    Else
        US_T_PS97 = T * 1.8 + 32
    End If
End Function
Function US_H_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute H_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2H97(P, S, H, Range)
    If Range = 0 Then
        US_H_PS97 = "Error!"
    Else
        US_H_PS97 = H / 2.326
    End If
End Function
Function US_V_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute V_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2V97(P, S, V, Range)
    If Range = 0 Then
        US_V_PS97 = "Error!"
    Else
        US_V_PS97 = V / 0.062428
    
End If
End Function
Function US_X_PS97(ByVal P As Double, ByVal S As Double)
Rem Attribute X_PS.VB_Description = "已知压力P(Psi)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_PS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, S As Double
    Rem P = 压力P
    Rem S = 比熵S
    P = 0.068948 * P / 10
    S = 4.1868 * S
   Call PS2X97(P, S, X, Range)
    If Range = 0 Then
        US_X_PS97 = "Error!"
    Else
        US_X_PS97 = X
    End If
End Function


Function US_T_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute T_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2T97(P, V, T, Range)
    If Range = 0 Then
        US_T_PV97 = "Error!"
    Else
        US_T_PV97 = T * 1.8 + 32
    End If
End Function
Function US_H_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute H_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2H97(P, V, H, Range)
    If Range = 0 Then
        US_H_PV97 = "Error!"
    Else
        US_H_PV97 = H / 2.326
    End If
End Function
Function US_S_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute S_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2S97(P, V, S, Range)
    If Range = 0 Then
        US_S_PV97 = "Error!"
    Else
        US_S_PV97 = S / 4.1868
    End If
End Function
Function US_X_PV97(ByVal P As Double, ByVal V As Double)
Rem Attribute X_PV.VB_Description = "已知压力P(Psi)和比容V(ft^3/lbm),求对应的干度X(100%)?"
Rem Attribute X_PV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim P As Double, V As Double
    Rem P = 压力P
    Rem V = 比容V
    P = 0.068948 * P / 10
    V = 0.062428 * V
    Call PV2X97(P, V, X, Range)
    If Range = 0 Then
        US_X_PV97 = "Error!"
    Else
        US_X_PV97 = X
    End If
End Function
Function US_T_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute T_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2T97(P, X, T, Range)
    If Range = 0 Then
        US_T_PX97 = "Error!"
    Else
        US_T_PX97 = T * 1.8 + 32
    End If
End Function
Function US_H_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute H_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2H97(P, X, H, Range)
    If Range = 0 Then
        US_H_PX97 = "Error!"
    Else
        US_H_PX97 = H / 2.326
    End If
End Function
Function US_S_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute S_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2S97(P, X, S, Range)
    If Range = 0 Then
        US_S_PX97 = "Error!"
    Else
        US_S_PX97 = S / 4.1868
    End If
End Function
Function US_V_PX97(ByVal P As Double, ByVal X As Double)
Rem Attribute V_PX.VB_Description = "已知压力P(Psi)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_PX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知P和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim P As Double, X As Double
    Rem P = 压力P
    Rem X = 干度X
    P = 0.068948 * P / 10
    Call PX2V97(P, X, V, Range)
    If Range = 0 Then
        US_V_PX97 = "Error!"
    Else
        US_V_PX97 = V / 0.062428
    
End If
End Function


Function US_P_T97(ByVal T As Double)
Rem Attribute P_T.VB_Description = "已知温度T(F),\r\n求对应的饱和压力P(Psi)?"
Rem Attribute P_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2P97(T, P, Range)
    If Range = 0 Then
        US_P_T97 = "Error!"
    Else
        US_P_T97 = P * 10 / 0.068948
    
End If
End Function
Function US_HL_T97(ByVal T As Double)
Rem Attribute Hw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比焓Hw(Btu/lbm)?"
Rem Attribute Hw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HW(饱和水焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2HL97(T, H, Range)
    If Range = 0 Then
        US_HL_T97 = "Error!"
    Else
        US_HL_T97 = H / 2.326
    End If
End Function
Function US_HG_T97(ByVal T As Double)
Rem Attribute Hs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比焓Hs(Btu/lbm)?"
Rem Attribute Hs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求HS(饱和蒸汽焓)。
    Dim H As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2HG97(T, H, Range)
    If Range = 0 Then
        US_HG_T97 = "Error!"
    Else
        US_HG_T97 = H / 2.326
    End If
End Function
Function US_SL_T97(ByVal T As Double)
Rem Attribute Ss_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比熵Ss( (Btu/lbmR) )?"
Rem Attribute Ss_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SS(饱和蒸汽熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SL97(T, S, Range)
    If Range = 0 Then
        US_SL_T97 = "Error!"
    Else
        US_SL_T97 = S / 4.1868
    End If
End Function
Function US_SG_T97(ByVal T As Double)
Rem Attribute Sw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比熵Sw( (Btu/lbmR) )?"
Rem Attribute Sw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求SW(饱和水熵)。
    Dim S As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SG97(T, S, Range)
    If Range = 0 Then
        US_SG_T97 = "Error!"
    Else
        US_SG_T97 = S / 4.1868
    End If
End Function
Function US_VL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2VL97(T, V, Range)
    If Range = 0 Then
        US_VL_T97 = "Error!"
    Else
        US_VL_T97 = V / 0.062428
    
End If
End Function
Function US_VG_T97(ByVal T As Double)
Rem Attribute Vs_T.VB_Description = "已知温度T(F),\r\n求对应的饱和蒸汽比容Vs(ft^3/lbm)?"
Rem Attribute Vs_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VS(饱和蒸汽比容)。
    Dim V As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2VG97(T, V, Range)
    If Range = 0 Then
        US_VG_T97 = "Error!"
    Else
        US_VG_T97 = V / 0.062428
    
End If
End Function


Function US_CpL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CPL97(T, CP, Range)
    If Range = 0 Then
        US_CpL_T97 = "Error!"
    Else
        US_CpL_T97 = CP
    End If
End Function


Function US_CpG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CPG97(T, CP, Range)
    If Range = 0 Then
        US_CpG_T97 = "Error!"
    Else
        US_CpG_T97 = CP
    End If
End Function


Function US_CvL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CVL97(T, CV, Range)
    If Range = 0 Then
        US_CvL_T97 = "Error!"
    Else
        US_CvL_T97 = CV
    End If
End Function



Function US_CvG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim CV As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2CVG97(T, CV, Range)
    If Range = 0 Then
        US_CvG_T97 = "Error!"
    Else
        US_CvG_T97 = CV
    End If
End Function

Function US_EL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EL97(T, E, Range)
    If Range = 0 Then
        US_EL_T97 = "Error!"
    Else
        US_EL_T97 = E
    End If
End Function


Function US_EG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim E As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EG97(T, E, Range)
    If Range = 0 Then
        US_EG_T97 = "Error!"
    Else
        US_EG_T97 = E
    End If
End Function


Function US_SSpL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SSPL97(T, SSP, Range)
    If Range = 0 Then
        US_SSpL_T97 = "Error!"
    Else
        US_SSpL_T97 = SSP
    End If
End Function


Function US_SSpG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SSP As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SSPG97(T, SSP, Range)
    If Range = 0 Then
        US_SSpG_T97 = "Error!"
    Else
        US_SSpG_T97 = SSP
    End If
End Function

Function US_KsL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2KSL97(T, KS, Range)
    If Range = 0 Then
        US_KsL_T97 = "Error!"
    Else
        US_KsL_T97 = KS
    End If
End Function

Function US_KsG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim KS As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2KSG97(T, KS, Range)
    If Range = 0 Then
        US_KsG_T97 = "Error!"
    Else
        US_KsG_T97 = KS
    End If
End Function

Function US_EtaL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2ETAL97(T, ETA, Range)
    If Range = 0 Then
        US_EtaL_T97 = "Error!"
    Else
        US_EtaL_T97 = ETA
    End If
End Function



Function US_EtaG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim ETA As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2ETAG97(T, ETA, Range)
    If Range = 0 Then
        US_EtaG_T97 = "Error!"
    Else
        US_EtaG_T97 = ETA
    End If
End Function


Function US_UL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2UL97(T, U, Range)
    If Range = 0 Then
        US_UL_T97 = "Error!"
    Else
        US_UL_T97 = U
    End If
End Function

Function US_UG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim U As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2UG97(T, U, Range)
    If Range = 0 Then
        US_UG_T97 = "Error!"
    Else
        US_UG_T97 = U
    End If
End Function

Function US_RamdL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2RAMDL97(T, RAMD, Range)
    If Range = 0 Then
        US_RamdL_T97 = "Error!"
    Else
        US_RamdL_T97 = RAMD
    End If
End Function


Function US_RamdG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim RAMD As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2RAMDG97(T, RAMD, Range)
    If Range = 0 Then
        US_RamdG_T97 = "Error!"
    Else
        US_RamdG_T97 = RAMD
    End If
End Function


Function US_EpsL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EPSL97(T, eps, Range)
    If Range = 0 Then
        US_EpsL_T97 = "Error!"
    Else
        US_EpsL_T97 = eps
    End If
End Function

Function US_EpsG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim eps As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2EPSG97(T, eps, Range)
    If Range = 0 Then
        US_EpsG_T97 = "Error!"
    Else
        US_EpsG_T97 = eps
    End If
End Function

Function US_PrnL_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2PRNL97(T, PRN, Range)
    If Range = 0 Then
        US_PrnL_T97 = "Error!"
    Else
        US_PrnL_T97 = PRN
    End If
End Function

Function US_PrnG_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim PRN As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2PRNG97(T, PRN, Range)
    If Range = 0 Then
        US_PrnG_T97 = "Error!"
    Else
        US_PrnG_T97 = PRN
    End If
End Function

Function US_NL_T97(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2NL97(T, Lamd, n, Range)
    If Range = 0 Then
        US_NL_T97 = "Error!"
    Else
        US_NL_T97 = n
    End If
End Function

Function US_NG_T97(ByVal T As Double, ByVal Lamd As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim n As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2NG97(T, Lamd, n, Range)
    If Range = 0 Then
        US_NG_T97 = "Error!"
    Else
        US_NG_T97 = n
    End If
End Function

Function US_SurfT_T97(ByVal T As Double)
Rem Attribute Vw_T.VB_Description = "已知温度T(F),\r\n求对应的饱和水比容Vw(ft^3/lbm)?"
Rem Attribute Vw_T.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T,求VW(饱和水比容)。
    Dim SurfT As Double, Range As Integer
    Rem Dim T As Double
    Rem T = 温度T
    T = (T - 32) / 1.8
    Call T2SURFT97(T, SurfT, Range)
    If Range = 0 Then
        US_SurfT_T97 = "Error!"
    Else
        US_SurfT_T97 = SurfT
    End If
End Function

Function US_P_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2P97(T, H, P, Range)
    If Range = 0 Then
        US_P_TH97 = "Error!"
    Else
        US_P_TH97 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2PLP97(T, H, P, Range)
    If Range = 0 Then
        US_PLP_TH97 = "Error!"
    Else
        US_PLP_TH97 = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute P_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2PHP97(T, H, P, Range)
    If Range = 0 Then
        US_PHP_TH97 = "Error!"
    Else
        US_PHP_TH97 = P * 10 / 0.068948
    
End If
End Function

Function US_S_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2S97(T, H, S, Range)
    If Range = 0 Then
        US_S_TH97 = "Error!"
    Else
        US_S_TH97 = S / 4.1868
    End If
End Function
Function US_SLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2SLP97(T, H, S, Range)
    If Range = 0 Then
        US_SLP_TH97 = "Error!"
    Else
        US_SLP_TH97 = S / 4.1868
    End If
End Function



Function US_SHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute S_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2SHP97(T, H, S, Range)
    If Range = 0 Then
        US_SHP_TH97 = "Error!"
    Else
        US_SHP_TH97 = S / 4.1868
    End If
End Function

Function US_V_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2V97(T, H, V, Range)
    If Range = 0 Then
        US_V_TH97 = "Error!"
    Else
        US_V_TH97 = V / 0.062428
    
End If
End Function
Function US_VLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2VLP97(T, H, V, Range)
    If Range = 0 Then
        US_VLP_TH97 = "Error!"
    Else
        US_VLP_TH97 = V / 0.062428
    
End If
End Function
Function US_VHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute V_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2VHP97(T, H, V, Range)
    If Range = 0 Then
        US_VHP_TH97 = "Error!"
    Else
        US_VHP_TH97 = V / 0.062428
    
End If
End Function

Function US_XLP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2XLP97(T, H, X, Range)
    If Range = 0 Then
        US_XLP_TH97 = "Error!"
    Else
        US_XLP_TH97 = X
    End If
End Function

Function US_XHP_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2XHP97(T, H, X, Range)
    If Range = 0 Then
        US_XHP_TH97 = "Error!"
    Else
        US_XHP_TH97 = X
    End If
End Function

Function US_X_TH97(ByVal T As Double, ByVal H As Double)
Rem Attribute X_TH.VB_Description = "已知温度T(F)和比焓H(Btu/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TH.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和H,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, H As Double
    Rem T = 温度T
    Rem H = 比焓H
    T = (T - 32) / 1.8
    H = 2.326 * H
    Call TH2X97(T, H, X, Range)
    If Range = 0 Then
        US_X_TH97 = "Error!"
    Else
        US_X_TH97 = X
    End If
End Function


Function US_P_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2P97(T, S, P, Range)
    If Range = 0 Then
        US_P_TS97 = "Error!"
    Else
        US_P_TS97 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2PLP97(T, S, P, Range)
    If Range = 0 Then
        US_PLP_TS97 = "Error!"
    Else
        US_PLP_TS97 = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute P_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2PHP97(T, S, P, Range)
    If Range = 0 Then
        US_PHP_TS97 = "Error!"
    Else
        US_PHP_TS97 = P * 10 / 0.068948
    
End If
End Function



Function US_H_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2H97(T, S, H, Range)
    If Range = 0 Then
        US_H_TS97 = "Error!"
    Else
        US_H_TS97 = H / 2.326
    End If
End Function


Function US_HLP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2HLP97(T, S, H, Range)
    If Range = 0 Then
        US_HLP_TS97 = "Error!"
    Else
        US_HLP_TS97 = H / 2.326
    End If
End Function


Function US_HHP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute H_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2HHP97(T, S, H, Range)
    If Range = 0 Then
        US_HHP_TS97 = "Error!"
    Else
        US_HHP_TS97 = H / 2.326
    End If
End Function




Function US_V_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2V97(T, S, V, Range)
    If Range = 0 Then
        US_V_TS97 = "Error!"
    Else
        US_V_TS97 = V / 0.062428
    
End If
End Function

Function US_VLP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2VLP97(T, S, V, Range)
    If Range = 0 Then
        US_VLP_TS97 = "Error!"
    Else
        US_VLP_TS97 = V / 0.062428
    
End If
End Function


Function US_VHP_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute V_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2VHP97(T, S, V, Range)
    If Range = 0 Then
        US_VHP_TS97 = "Error!"
    Else
        US_VHP_TS97 = V / 0.062428
    
End If
End Function


Function US_X_TS97(ByVal T As Double, ByVal S As Double)
Rem Attribute X_TS.VB_Description = "已知温度T(F)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_TS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, S As Double
    Rem T = 温度T
    Rem S = 比熵S
    T = (T - 32) / 1.8
    S = 4.1868 * S
    Call TS2X97(T, S, X, Range)
    If Range = 0 Then
        US_X_TS97 = "Error!"
    Else
        US_X_TS97 = X
    End If
End Function
Function US_P_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute P_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2P97(T, V, P, Range)
    If Range = 0 Then
        US_P_TV97 = "Error!"
    Else
        US_P_TV97 = P * 10 / 0.068948
    
End If
End Function
Function US_H_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute H_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2H97(T, V, H, Range)
    If Range = 0 Then
        US_H_TV97 = "Error!"
    Else
        US_H_TV97 = H / 2.326
    End If
End Function
Function US_S_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute S_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2S97(T, V, S, Range)
    If Range = 0 Then
        US_S_TV97 = "Error!"
    Else
        US_S_TV97 = S / 4.1868
    End If
End Function
Function US_X_TV97(ByVal T As Double, ByVal V As Double)
Rem Attribute X_TV.VB_Description = "已知温度T(F)和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_TV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim T As Double, V As Double
    Rem T = 温度T
    Rem V = 比容V
    T = (T - 32) / 1.8
    V = 0.062428 * V
    Call TV2X97(T, V, X, Range)
    If Range = 0 Then
        US_X_TV97 = "Error!"
    Else
        US_X_TV97 = X
    End If
End Function
Function US_P_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute P_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2P97(T, X, P, Range)
    If Range = 0 Then
        US_P_TX97 = "Error!"
    Else
        US_P_TX97 = P * 10 / 0.068948
    
End If
End Function
Function US_H_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute H_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2H97(T, X, H, Range)
    If Range = 0 Then
        US_H_TX97 = "Error!"
    Else
        US_H_TX97 = H / 2.326
    End If
End Function
Function US_S_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute S_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2S97(T, X, S, Range)
    If Range = 0 Then
        US_S_TX97 = "Error!"
    Else
        US_S_TX97 = S / 4.1868
    End If
End Function
Function US_V_TX97(ByVal T As Double, ByVal X As Double)
Rem Attribute V_TX.VB_Description = "已知温度T(F)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_TX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知T和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim T As Double, X As Double
    Rem T = 温度T
    Rem X = 干度X
    T = (T - 32) / 1.8
    Call TX2V97(T, X, V, Range)
    If Range = 0 Then
        US_V_TX97 = "Error!"
    Else
        US_V_TX97 = V / 0.062428
    
End If
End Function


Function US_P_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute P_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2P97(H, S, P, Range)
    If Range = 0 Then
        US_P_HS97 = "Error!"
    Else
        US_P_HS97 = P * 10 / 0.068948
    
End If
End Function

Function US_T_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute T_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的温度T(F)?"
Rem Attribute T_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2T97(H, S, T, Range)
    If Range = 0 Then
        US_T_HS97 = "Error!"
    Else
        US_T_HS97 = T * 1.8 + 32
    End If
End Function

Function US_V_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute V_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2V97(H, S, V, Range)
    If Range = 0 Then
        US_V_HS97 = "Error!"
    Else
        US_V_HS97 = V / 0.062428
    
End If
End Function

Function US_X_HS97(ByVal H As Double, ByVal S As Double)
Rem Attribute X_HS.VB_Description = "已知比焓H(Btu/lbm)和比熵S( (Btu/lbmR) ),\r\n求对应的干度X(100%)?"
Rem Attribute X_HS.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和S,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, S As Double
    Rem H = 比焓H
    Rem S = 比熵S
    H = 2.326 * H
    S = 4.1868 * S
    Call HS2X97(H, S, X, Range)
    If Range = 0 Then
        US_X_HS97 = "Error!"
    Else
        US_X_HS97 = X
    End If
End Function

Function US_P_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute P_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2P97(H, V, P, Range)
    If Range = 0 Then
        US_P_HV97 = "Error!"
    Else
        US_P_HV97 = P * 10 / 0.068948
    
End If
End Function

Function US_T_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute T_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2T97(H, V, T, Range)
    If Range = 0 Then
        US_T_HV97 = "Error!"
    Else
        US_T_HV97 = T * 1.8 + 32
    End If
End Function

Function US_S_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute S_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2S97(H, V, S, Range)
    If Range = 0 Then
        US_S_HV97 = "Error!"
    Else
        US_S_HV97 = S / 4.1868
    End If
End Function

Function US_X_HV97(ByVal H As Double, ByVal V As Double)
Rem Attribute X_HV.VB_Description = "已知比焓H(Btu/lbm)和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_HV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim H As Double, V As Double
    Rem H = 比焓H
    Rem V = 比容V
    H = 2.326 * H
    V = 0.062428 * V
    Call HV2X97(H, V, X, Range)
    If Range = 0 Then
        US_X_HV97 = "Error!"
    Else
        US_X_HV97 = X
    End If
End Function

Function US_P_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2P97(H, X, P, Range)
    If Range = 0 Then
        US_P_HX97 = "Error!"
    Else
        US_P_HX97 = P * 10 / 0.068948
    
End If
End Function

Function US_PLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2PLP97(H, X, P, Range)
    If Range = 0 Then
        US_PLP_HX97 = "Error!"
    Else
        US_PLP_HX97 = P * 10 / 0.068948
    
End If
End Function


Function US_PHP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute P_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2PHP97(H, X, P, Range)
    If Range = 0 Then
        US_PHP_HX97 = "Error!"
    Else
        US_PHP_HX97 = P * 10 / 0.068948
    
End If
End Function


Function US_T_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2T97(H, X, T, Range)
    If Range = 0 Then
        US_T_HX97 = "Error!"
    Else
        US_T_HX97 = T * 1.8 + 32
    End If
End Function


Function US_TLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2TLP97(H, X, T, Range)
    If Range = 0 Then
        US_TLP_HX97 = "Error!"
    Else
        US_TLP_HX97 = T * 1.8 + 32
    End If
End Function


Function US_THP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute T_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2THP97(H, X, T, Range)
    If Range = 0 Then
        US_THP_HX97 = "Error!"
    Else
        US_THP_HX97 = T * 1.8 + 32
    End If
End Function

Function US_S_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2S97(H, X, S, Range)
    If Range = 0 Then
        US_S_HX97 = "Error!"
    Else
        US_S_HX97 = S / 4.1868
    End If
End Function

Function US_SLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2SLP97(H, X, S, Range)
    If Range = 0 Then
        US_SLP_HX97 = "Error!"
    Else
        US_SLP_HX97 = S / 4.1868
    End If
End Function

Function US_SHP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute S_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2SHP97(H, X, S, Range)
    If Range = 0 Then
        US_SHP_HX97 = "Error!"
    Else
        US_SHP_HX97 = S / 4.1868
    End If
End Function

Function US_V_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2V97(H, X, V, Range)
    If Range = 0 Then
        US_V_HX97 = "Error!"
    Else
        US_V_HX97 = V / 0.062428
    
End If
End Function

Function US_VLP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2VLP97(H, X, V, Range)
    If Range = 0 Then
        US_VLP_HX97 = "Error!"
    Else
        US_VLP_HX97 = V / 0.062428
    
End If
End Function

Function US_VHP_HX97(ByVal H As Double, ByVal X As Double)
Rem Attribute V_HX.VB_Description = "已知比焓H(Btu/lbm)和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_HX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知H和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem H = 比焓H
    Rem X = 干度X
    H = 2.326 * H
    Call HX2VHP97(H, X, V, Range)
    If Range = 0 Then
        US_VHP_HX97 = "Error!"
    Else
        US_VHP_HX97 = V / 0.062428
    
End If
End Function


Function US_P_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute P_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求P。
    Dim P As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2P97(S, V, P, Range)
    If Range = 0 Then
        US_P_SV97 = "Error!"
    Else
        US_P_SV97 = P * 10 / 0.068948
    
End If
End Function

Function US_T_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute T_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的温度T(F)?"
Rem Attribute T_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求T。
    Dim T As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2T97(S, V, T, Range)
    If Range = 0 Then
        US_T_SV97 = "Error!"
    Else
        US_T_SV97 = T * 1.8 + 32
    End If
End Function

Function US_H_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute H_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求H。
    Dim H As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2H97(S, V, H, Range)
    If Range = 0 Then
        US_H_SV97 = "Error!"
    Else
        US_H_SV97 = H / 2.326
    End If
End Function

Function US_X_SV97(ByVal S As Double, ByVal V As Double)
Rem Attribute X_SV.VB_Description = "已知比熵S( (Btu/lbmR) )和比容V(ft^3/lbm),\r\n求对应的干度X(100%)?"
Rem Attribute X_SV.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和V,求X。
    Dim X As Double, Range As Integer
    Rem Dim S As Double, V As Double
    Rem S = 比熵S
    Rem V = 比容V
    S = 4.1868 * S
    V = 0.062428 * V
    Call SV2X97(S, V, X, Range)
    If Range = 0 Then
        US_X_SV97 = "Error!"
    Else
        US_X_SV97 = X
    End If
End Function

Function US_P_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2P97(S, X, P, Range)
    If Range = 0 Then
        US_P_SX97 = "Error!"
    Else
        US_P_SX97 = P * 10 / 0.068948
    
End If
End Function


Function US_PLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PLP97(S, X, P, Range)
    If Range = 0 Then
        US_PLP_SX97 = "Error!"
    Else
        US_PLP_SX97 = P * 10 / 0.068948
    
End If
End Function

Function US_PMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PMP97(S, X, P, Range)
    If Range = 0 Then
        US_PMP_SX97 = "Error!"
    Else
        US_PMP_SX97 = P * 10 / 0.068948
    
End If
End Function

Function US_PHP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute P_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2PHP97(S, X, P, Range)
    If Range = 0 Then
        US_PHP_SX97 = "Error!"
    Else
        US_PHP_SX97 = P * 10 / 0.068948
    
End If
End Function
Function US_T_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2T97(S, X, T, Range)
    If Range = 0 Then
        US_T_SX97 = "Error!"
    Else
        US_T_SX97 = T * 1.8 + 32
    End If
End Function

Function US_TLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2TLP97(S, X, T, Range)
    If Range = 0 Then
        US_TLP_SX97 = "Error!"
    Else
        US_TLP_SX97 = T * 1.8 + 32
    End If
End Function

Function US_TMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2TMP97(S, X, T, Range)
    If Range = 0 Then
        US_TMP_SX97 = "Error!"
    Else
        US_TMP_SX97 = T * 1.8 + 32
    End If
End Function

Function US_THP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute T_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2THP97(S, X, T, Range)
    If Range = 0 Then
        US_THP_SX97 = "Error!"
    Else
        US_THP_SX97 = T * 1.8 + 32
    End If
End Function

Function US_H_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2H97(S, X, H, Range)
    If Range = 0 Then
        US_H_SX97 = "Error!"
    Else
        US_H_SX97 = H / 2.326
    End If
End Function

Function US_HLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HLP97(S, X, H, Range)
    If Range = 0 Then
        US_HLP_SX97 = "Error!"
    Else
        US_HLP_SX97 = H / 2.326
    End If
End Function

Function US_HMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HMP97(S, X, H, Range)
    If Range = 0 Then
        US_HMP_SX97 = "Error!"
    Else
        US_HMP_SX97 = H / 2.326
    End If
End Function

Function US_HHP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute H_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2HHP97(S, X, H, Range)
    If Range = 0 Then
        US_HHP_SX97 = "Error!"
    Else
        US_HHP_SX97 = H / 2.326
    End If
End Function

Function US_V_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2V97(S, X, V, Range)
    If Range = 0 Then
        US_V_SX97 = "Error!"
    Else
        US_V_SX97 = V / 0.062428
    
End If
End Function

Function US_VLP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VLP97(S, X, V, Range)
    If Range = 0 Then
        US_VLP_SX97 = "Error!"
    Else
        US_VLP_SX97 = V / 0.062428
    
End If
End Function

Function US_VMP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VMP97(S, X, V, Range)
    If Range = 0 Then
        US_VMP_SX97 = "Error!"
    Else
        US_VMP_SX97 = V / 0.062428
    
End If
End Function

Function US_VHP_SX97(ByVal S As Double, ByVal X As Double)
Rem Attribute V_SX.VB_Description = "已知比熵S( (Btu/lbmR) )和干度X(100%),\r\n求对应的比容V(ft^3/lbm)?"
Rem Attribute V_SX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知S和X,求V。
    Dim V As Double, Range As Integer
    Rem Dim H As Double, X As Double
    Rem S = 比熵S
    Rem X = 干度X
    S = 4.1868 * S
    Call SX2VHP97(S, X, V, Range)
    If Range = 0 Then
        US_VHP_SX97 = "Error!"
    Else
        US_VHP_SX97 = V / 0.062428
    
End If
End Function

Function US_P_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2P97(V, X, P, Range)
    If Range = 0 Then
        US_P_VX97 = "Error!"
    Else
        US_P_VX97 = P * 10 / 0.068948
    
End If
End Function
Function US_PLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2PLP97(V, X, P, Range)
    If Range = 0 Then
        US_PLP_VX97 = "Error!"
    Else
        US_PLP_VX97 = P * 10 / 0.068948
    
End If
End Function
Function US_PHP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute P_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的压力P(Psi)?"
Rem Attribute P_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求P。
    Dim P As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2PHP97(V, X, P, Range)
    If Range = 0 Then
        US_PHP_VX97 = "Error!"
    Else
        US_PHP_VX97 = P * 10 / 0.068948
    
End If
End Function

Function US_T_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2T97(V, X, T, Range)
    If Range = 0 Then
        US_T_VX97 = "Error!"
    Else
        US_T_VX97 = T * 1.8 + 32
    End If
End Function

Function US_TLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2TLP97(V, X, T, Range)
    If Range = 0 Then
        US_TLP_VX97 = "Error!"
    Else
        US_TLP_VX97 = T * 1.8 + 32
    End If
End Function

Function US_THP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute T_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的温度T(F)?"
Rem Attribute T_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求T。
    Dim T As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2THP97(V, X, T, Range)
    If Range = 0 Then
        US_THP_VX97 = "Error!"
    Else
        US_THP_VX97 = T * 1.8 + 32
    End If
End Function

Function US_H_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2H97(V, X, H, Range)
    If Range = 0 Then
        US_H_VX97 = "Error!"
    Else
        US_H_VX97 = H / 2.326
    End If
End Function

Function US_HLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2HLP97(V, X, H, Range)
    If Range = 0 Then
        US_HLP_VX97 = "Error!"
    Else
        US_HLP_VX97 = H / 2.326
    End If
End Function

Function US_HHP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute H_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比焓H(Btu/lbm)?"
Rem Attribute H_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求H。
    Dim H As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2HHP97(V, X, H, Range)
    If Range = 0 Then
        US_HHP_VX97 = "Error!"
    Else
        US_HHP_VX97 = H / 2.326
    End If
End Function

Function US_S_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2S97(V, X, S, Range)
    If Range = 0 Then
        US_S_VX97 = "Error!"
    Else
        US_S_VX97 = S / 4.1868
    End If
End Function

Function US_SLP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2SLP97(V, X, S, Range)
    If Range = 0 Then
        US_SLP_VX97 = "Error!"
    Else
        US_SLP_VX97 = S / 4.1868
    End If
End Function

Function US_SHP_VX97(ByVal V As Double, ByVal X As Double)
Rem Attribute S_VX.VB_Description = "已知比容V(ft^3/lbm)和干度X(100%),\r\n求对应的比熵S( (Btu/lbmR) )?"
Rem Attribute S_VX.VB_ProcData.VB_Invoke_Func = " \n16"
Rem 已知V和X,求S。
    Dim S As Double, Range As Integer
    Rem Dim V As Double, X As Double
    Rem V = 比容V
    Rem X = 干度X
    V = 0.062428 * V
    Call VX2SHP97(V, X, S, Range)
    If Range = 0 Then
        US_SHP_VX97 = "Error!"
    Else
        US_SHP_VX97 = S / 4.1868
    End If
End Function

Rem 两点线性插值
Rem function INT2DXX(ByVal X1 As Double, ByVal X2 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal x As Double) As Double
Rem Attribute IN2DX_Y.VB_Description = "根据两点线性插值"
Rem Attribute IN2DX_Y.VB_ProcData.VB_Invoke_Func = " \n16"
Rem    Dim y As Double
Rem    Call INST2DXX(X1, X2, Y1, Y2, x, y)
Rem    INT2DXX = y
Rem End Function


Rem function INT2DXY(ByVal X1 As Double, ByVal Y1 As Double, ByVal X2 As Double, ByVal Y2 As Double, ByVal x As Double) As Double
Rem Attribute P2_XY.VB_Description = "根据两点线性插值"
Rem Attribute P2_XY.VB_ProcData.VB_Invoke_Func = " \n16"
Rem    Dim y As Double
Rem    Call INST2DXY(X1, Y1, X2, Y2, x, y)
Rem    INT2DXY = y
Rem End Function

Public Function my_GETSTD_WASP()
Dim std As Integer
  Call GETSTD_WASP(std)
  my_GETSTD_WASP = std
End Function


Public Sub my_SETSTD_WASP(ByVal std As Integer)
  Call SETSTD_WASP(std)
End Sub
