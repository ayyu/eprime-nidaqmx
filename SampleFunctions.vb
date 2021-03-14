'My own wrapper functions for controlling the NI DAQ device

Sub CreateAOVoltageChan(DAQtaskHandle As Long, Channel As String)
	DAQstatus = DAQmxCreateAOVoltageChan(DAQtaskHandle, Channel, "", -10.0, 10.0, DAQmx_Val_Volts, "")
End Sub

Sub CreateDOChan(DAQtaskHandle As Long, Channel As String)
	DAQstatus = DAQmxCreateDOChan(DAQtaskHandle, Channel, "", DAQmx_Val_ChanPerLine)
End Sub

Sub InitDAQTask(DAQtaskHandle As Long)
	DAQstatus = DAQmxCreateTask("", DAQtaskHandle)
	DAQstatus = DAQmxStartTask(DAQtaskHandle)
End Sub

Sub WriteAOVoltage(DAQtaskHandle As Long, Value As Double)
	Dim OutputVoltage(0) As Double
	OutputVoltage(0) = Value
	DAQstatus = DAQmxWriteAnalogF64(DAQtaskHandle, 1, True, DAQmx_Val_WaitInfinitely, DAQmx_Val_GroupByChannel, OutputVoltage, DAQsampsPerChanWritten, ByVal 0&)
End Sub

Sub PulseAOVoltage(DAQtaskHandle As Long, Value As Double, NumSamples As Integer)
	Dim OutputVoltage(NumSamples - 1) As Double
	Dim i As Long
	For i = 0 To (NumSamples - 2)
		OutputVoltage(i) = Value
	Next
	DAQstatus = DAQmxWriteAnalogF64(DAQtaskHandle, NumSamples, True, DAQmx_Val_WaitInfinitely, DAQmx_Val_GroupByChannel, OutputVoltage, DAQsampsPerChanWritten, ByVal 0&)
End Sub

Sub WriteDOLines(DAQtaskHandle As Long, Value As Integer)
	Dim OutputTrain(0) As Integer
	OutputTrain(0) = Value
	DAQstatus = DAQmxWriteDigitalLines(DAQtaskHandle, 1, True, DAQmx_Val_WaitInfinitely, DAQmx_Val_GroupByChannel, OutputTrain, DAQsampsPerChanWritten, ByVal 0&)
End Sub

Sub PulseDOLines(DAQtaskHandle As Long, Value As Integer, NumSamples As Integer)
	Dim OutputTrain(NumSamples - 1) As Integer
	Dim i As Long
	For i = 0 To (NumSamples - 2)
		OutputTrain(i) = Value
	Next
	DAQstatus = DAQmxWriteDigitalLines(DAQtaskHandle, NumSamples, True, DAQmx_Val_WaitInfinitely, DAQmx_Val_GroupByChannel, OutputTrain, DAQsampsPerChanWritten, ByVal 0&)
End Sub

Sub EndDAQTask(DAQtaskHandle As Long)
	DAQstatus = DAQmxStopTask(DAQtaskHandle)
	DAQstatus = DAQmxClearTask(DAQtaskHandle)
End Sub