Imports Microsoft
Imports System

Module MDL_Delta

#Region "Variable Definition"

#Region "Communication"

    Public MsCom_Mnl_Tst() As System.IO.Ports.SerialPort

#End Region

#Region "Flow Calibration"

    Public Flow_01_Gain As Single = 1
    Public Flow_02_Gain As Single = 1
    Public Flow_03_Gain As Single = 1
    Public Flow_04_Gain As Single = 1
    Public Flow_05_Gain As Single = 1

    Public Flow_01_Offset As Single = 0
    Public Flow_02_Offset As Single = 0
    Public Flow_03_Offset As Single = 0
    Public Flow_04_Offset As Single = 0
    Public Flow_05_Offset As Single = 0

    Public Pump_01_Eff_Coe As Single = 0.01
    Public Pump_02_Eff_Coe As Single = 0.01
    Public Pump_03_Eff_Coe As Single = 0.01
    Public Pump_04_Eff_Coe As Single = 0.01
    Public Pump_05_Eff_Coe As Single = 0.01

#End Region

#Region "PLC_Variables"

#Region "Constant"

    Public Const PLC_Delta_D_I_Const = &H800
    Public Const PLC_Delta_D_O_Const = &H810
    Public Const PLC_Coil_On = &HFF00
    Public Const PLC_Delta_Coil_Off = &H0

#End Region

#Region "Enums"

    Public Enum D_I
        Pannel_Emerg = &H0
        Outdoor_Emerg = &H1
        Table_Emerg = &H2
        Air_Pressure_SW = &H3
        Petrol_Level_SW = &H4
        Water_Level_SW = &H5
        Water_Termostate = &H6
    End Enum

    Public Enum D_O
        I2P_En_01 = &H10
        I2P_En_02 = &H11
        I2P_En_03 = &H12
        I2P_En_04 = &H13
        I2P_En_05 = &H14
        Heater_En = &H15
        Water_Pump_En = &H16
        Blender_En = &H17
    End Enum

    Public Enum PLC_Err_Code
        Illegal_Command = &H1
        Illegal_Device_Add = &H2
        Illegal_Device_Val = &H3
        Check_Sum_Err = &H7
    End Enum

    Public Enum PLC_Command_Code
        Read_Coil_Status = &H1
        Read_Input_Status = &H2
        Read_Holding_Reg = &H3
        Write_Single_Coil = &H5
        Write_Single_Reg = &H6
        Write_Multiple_Coils = &HF
        Write_Multiple_reg = &H10
        PLC_Status = &H11
    End Enum

    Public Enum PLC_Device_Offset
        S = &H0
        X = &H400
        Y = &H500
        T = &H600
        M = &H800
        C = &HE00
        D = &H1000
    End Enum

#End Region

#End Region

#Region "Command Enum"

    Public Enum Delta_Command
        Read_Bits = &H2
        Read_Register_s = &H3
        Write_1_Bit = &H5
        Write_1_Word = &H6
    End Enum

#End Region

#Region "Error Enum"

    Public Enum Delta_Read_Error
        Initial_Proce_Temperature_Value_Is_Not_Got_Yet = -32766
        Temperature_Sensor_Is_Not_Connected = -32765
        Temperature_Sensor_Input_Error = -32764
        ADC_Input_Error = -32763
        Memory_Read_Write_Err = -32762
    End Enum

    Public Enum Delta_Com_Ack_Error
        Preset_Value_Unstable = 1
        Re_Initial_No_Temperature_At_this_Time = 2
        Input_Sensor_Did_Not_Connect = 3
        Input_Signal_Error = 4
        Over_Input_Range = 5
        ADC_Fail = 6
        EEPROM_Read_Write_Error = 7
    End Enum

#End Region

#Region " Reg/Bit Address"

    Public Enum Delta_Reg_Add

        'First 16 Segment Range
        PV_Process_Value = &H1000
        SV_Set_Point = &H1001
        Upper_Limit_Of_Temperature_Range = &H1002
        Lower_Limit_Of_Temperature_Range = &H1003
        Input_Temperature_Sensor_Type = &H1004
        Control_Method = &H1005
        Heating_Colling_Control_Selection = &H1006
        Fst_Group_Of_Heating_Cooling_Control_Cycle = &H1007
        Snd_Group_Of_Heating_Cooling_Control_Cycle = &H1008
        PB_Proportional_Band = &H1009
        Ti_Integral_Time = &H100A
        Td_Derivative_Time = &H100B
        Integration_Default = &H100C
        Proportional_Control_Offset_Error_Value_When_Ti_is_0 = &H100D
        The_Setting_Of_COEF_When_Dual_Loop_Output_Control_Are_Used = &H100E
        The_Setting_Of_Dead_Band_When_Dual_Loop_Output_Control_Are_Used = &H100F

        'Second 16 Segment Range
        Hystersis_Setting_Value_Of_The_1st_Output_Group = &H1010
        Hystersis_Setting_Value_Of_The_2nd_Output_Group = &H1011
        Output_Value_Read_And_Write_Of_Output_1 = &H1012
        Output_Value_Read_And_Write_Of_Output_2 = &H1013
        Upper_Limit_Regulation_Of_Analog_Linear_Output = &H1014
        Lower_Limit_Regulation_Of_Analog_Linear_Output = &H1015
        Temperature_Regulation_Value = &H1016
        Analog_Decimal_Setting = &H1017
        Time_For_Valve_From_Full_Open_To_Full_Close = &H1018
        Dead_Band_Setting_Of_Vavle = &H1019
        Upper_Limit_Of_FeedBack_Signal_Set_By_Vavle = &H101A
        Lower_Limit_Of_FeedBack_Signal_Set_Bye_Valve = &H101B
        PID_Parameter_selection = &H101C
        Set_Value_Coresponded_to_PID_Value = &H101D

        'Third 16 Segment Range
        Alarm_1_Type = &H1020
        Alarm_2_Type = &H1021
        Alarm_3_Type = &H1022
        System_Alarm_Setting = &H1023
        Upper_Limit_Alarm_1 = &H1024
        Lower_Limit_Alarm_1 = &H1025
        Upper_Limit_Alarm_2 = &H1026
        Lower_Limit_Alarm_2 = &H1027
        Upper_Limit_Alarm_3 = &H1028
        Lower_Limit_Alarm_3 = &H1029
        Red_LED_Status = &H102A
        Read_Pushbutton_Status = &H102B
        Setting_Lock_Status = &H102C
        CT_read_Value = &H102D
        Software_Version = &H102F

        'Forth 16 Segment Range
        Start_Pattern_Number = &H1030

        'Fifth 16 Segment Range
        Actual_Step_Number_Setting_Inside_Pattern_0 = &H1040
        Actual_Step_Number_Setting_Inside_Pattern_1 = &H1041
        Actual_Step_Number_Setting_Inside_Pattern_2 = &H1042
        Actual_Step_Number_Setting_Inside_Pattern_3 = &H1043
        Actual_Step_Number_Setting_Inside_Pattern_4 = &H1044
        Actual_Step_Number_Setting_Inside_Pattern_5 = &H1045
        Actual_Step_Number_Setting_Inside_Pattern_6 = &H1046
        Actual_Step_Number_Setting_Inside_Pattern_7 = &H1047

        'Sixth 16 Segmet Range
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_0 = &H1050
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_1 = &H1051
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_2 = &H1052
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_3 = &H1053
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_4 = &H1054
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_5 = &H1055
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_6 = &H1056
        Cycle_Number_For_Repeating_The_Execution_of_Pattern_7 = &H1057

    End Enum

    Public Enum Delta_Bit_Add

        Communication_Write_In_Selection = &H810
        Temperature_Unit_Display_Selection = &H811
        Decimal_Point_Position_Selection = &H812
        AT_Setting = &H813
        Control_Run_Stop_Setting = &H814
        Stop_Setting_For_PID_Program_Control = &H815
        Temporarily_Stop_For_PID_Program_Control = &H816
        Valve_Feedback_Setting_Status = &H817
        Auto_Tuning_Valve_Feedback_Status = &H818

    End Enum

#End Region

#Region "Reg Contents"

    Public Enum Control_Method
        PID = 0
        ON_OFF = 1
        Manual_Tuning = 2
        Cooling_Heating = 3
    End Enum

    Public Enum Heating_Cooling_Control_Selection
        Heating = 0
        Cooling = 1
        Heating_Cooling = 2
        Cooling_heating = 3
    End Enum

    Public Enum System_Alarm_Setting
        None = 0
        Alarm_01 = 1
        Alarm_02 = 2
        Alarm_03 = 3
    End Enum

    Public Enum Setting_Lick_Status
        Normal = 0
        All_Setting_Lick = 1
        Lock_Others_Than_SV_Value = 11
    End Enum

#End Region

#End Region

#Region "Structures"

#End Region

#Region "Functions Definitions"

    Public Function Set_LRC(ByVal Str1 As String) As String
        Dim LCR As Integer
        Dim I As Integer
        Dim C As String

        LCR = 0
        For I = 1 To Len(Str1) Step 2
            LCR = LCR + ConvertHexToInteger(Mid(Str1, I, 2))
        Next I
        If LCR < 255 Then
            C = Trim(Hex(255 - LCR + 1))
        Else
            C = Trim(Hex(65535 - LCR + 1))
        End If
        If Len(Trim(C)) > 2 Then C = Right$(C, 2)
        Select Case Len(Trim(C))
            Case 0 : C = "00"
            Case 1 : C = "0" & Trim(C)
            Case 2 : C = Trim(C)
            Case Else : C = Format(Trim(C), "00")
        End Select
        Set_LRC = Str1 & Trim(C)
    End Function

    Public Function DataRecieve_PreFunc(ByVal Delta_Add As Integer, ByVal Reg_Add As Integer) As String
        Try
            Dim Str_Command As String = "03"
            Dim Str_Delta_Add As String = Mid(ConvertIntegerToHex(Delta_Add), 5, 4)
            Str_Delta_Add = Mid(Str_Delta_Add, Str_Delta_Add.Length - 1, 2)
            Dim Str_Reg_Add As String = Mid(ConvertIntegerToHex(Reg_Add), 5, 4)
            Dim Str_Read_No As String = "0001"
            DataRecieve_PreFunc = Str_Delta_Add & Str_Command & Str_Reg_Add & Str_Read_No
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
            DataRecieve_PreFunc = ""
        End Try
    End Function

    Public Function DatRecieve_Command(ByVal Str_Recieve As String) As String
        Dim Str_Command As String = Mid(Str_Recieve, 4, 2)
        Select Case Str_Command
            Case "02"
                Return "Bit Reading"
            Case "03"
                Return "Rgister Reading"
            Case "06"
                Return "Word Writing"
            Case Else
                Return "Unknow Command!"
        End Select

    End Function

    Public Function DataRecieve_Value(ByVal Str_Recieve As String) As Integer
        Dim Str_Command As String = Mid(Str_Recieve, 4, 2)
        DataRecieve_Value = ConvertHexToInteger(Mid(Str_Recieve, 8, 4))
    End Function

    Public Function Flow_Calibration_Read() As Boolean
        Dim FileName As String = "C:\WINDOWS\system32\Calibration.Ini"
        Dim FileInput As IO.FileStream = Nothing
        Dim BinaryReader As IO.BinaryReader = Nothing
        Try
            If Not System.IO.File.Exists(FileName) Then
                Call Flow_Calibration_Write()
            End If
            FileInput = New IO.FileStream(FileName, IO.FileMode.Open, IO.FileAccess.Read)
            BinaryReader = New IO.BinaryReader(FileInput)
            FileInput.Seek(0, IO.SeekOrigin.Begin)
            Flow_01_Gain = BinaryReader.ReadSingle
            Flow_02_Gain = BinaryReader.ReadSingle
            Flow_03_Gain = BinaryReader.ReadSingle
            Flow_04_Gain = BinaryReader.ReadSingle
            Flow_05_Gain = BinaryReader.ReadSingle

            Flow_01_Offset = BinaryReader.ReadSingle
            Flow_02_Offset = BinaryReader.ReadSingle
            Flow_03_Offset = BinaryReader.ReadSingle
            Flow_04_Offset = BinaryReader.ReadSingle
            Flow_05_Offset = BinaryReader.ReadSingle

            Pump_01_Eff_Coe = BinaryReader.ReadSingle
            Pump_02_Eff_Coe = BinaryReader.ReadSingle
            Pump_03_Eff_Coe = BinaryReader.ReadSingle
            Pump_04_Eff_Coe = BinaryReader.ReadSingle
            Pump_05_Eff_Coe = BinaryReader.ReadSingle

            If Not FileInput Is Nothing Then FileInput.Close()
            If BinaryReader Is Nothing Then BinaryReader.Close()

            Return True
        Catch ex As Exception
            If Not FileInput Is Nothing Then FileInput.Close()
            If BinaryReader Is Nothing Then BinaryReader.Close()
            Return False
        End Try
    End Function

    Public Function Flow_Calibration_Write() As Boolean
        Dim FileName As String = "C:\WINDOWS\system32\Calibration.Ini"
        Dim FileOutput As IO.FileStream = Nothing
        Dim BinaryOutput As IO.BinaryWriter = Nothing
        Try
            FileOutput = New IO.FileStream(FileName, IO.FileMode.Create, IO.FileAccess.Write)
            FileOutput.SetLength(40)
            BinaryOutput = New IO.BinaryWriter(FileOutput)
            FileOutput.Position = 0

            BinaryOutput.Write(Flow_01_Gain)
            BinaryOutput.Write(Flow_02_Gain)
            BinaryOutput.Write(Flow_03_Gain)
            BinaryOutput.Write(Flow_04_Gain)
            BinaryOutput.Write(Flow_05_Gain)

            BinaryOutput.Write(Flow_01_Offset)
            BinaryOutput.Write(Flow_02_Offset)
            BinaryOutput.Write(Flow_03_Offset)
            BinaryOutput.Write(Flow_04_Offset)
            BinaryOutput.Write(Flow_05_Offset)

            BinaryOutput.Write(Pump_01_Eff_Coe)
            BinaryOutput.Write(Pump_02_Eff_Coe)
            BinaryOutput.Write(Pump_03_Eff_Coe)
            BinaryOutput.Write(Pump_04_Eff_Coe)
            BinaryOutput.Write(Pump_05_Eff_Coe)

            If Not FileOutput Is Nothing Then FileOutput.Close()
            If BinaryOutput Is Nothing Then BinaryOutput.Close()
            Return True
        Catch ex As Exception
            If Not FileOutput Is Nothing Then FileOutput.Close()
            If BinaryOutput Is Nothing Then BinaryOutput.Close()
            Return False
        End Try
    End Function

#End Region

End Module
