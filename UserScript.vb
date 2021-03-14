'Some variables to store return values from the DLL functions

Dim DAQstatus As Long
Dim DAQsampsPerChanWritten As Long

'See https://nidaqmx-python.readthedocs.io/en/latest/constants.html

Const DAQmx_Val_Volts As Long = 10348
Const DAQmx_Val_GroupByScanNumber As Long = 1 
Const DAQmx_Val_GroupByChannel As Long = 0
Const DAQmx_Val_WaitInfinitely As Double = -1
Const DAQmx_Val_ChanPerLine As Long = 0
Const DAQmx_Val_ChanForAllLines As Long = 1

Const DAQmx_Val_Rising As Long = 10280
Const DAQmx_Val_Falling As Long =  10171
Const DAQmx_Val_FiniteSamps As Long = 10178
Const DAQmx_Val_ContSamps As Long = 10123

'Expose variables in the NI DAQmx DLL

Declare Function DAQmxResetDevice Lib "nicaiu.dll" (ByVal deviceName As String) As Long

Declare Function DAQmxCreateTask Lib "nicaiu.dll" (ByVal staskName As String, ByRef taskHandle As Long) As Long
Declare Function DAQmxStartTask Lib "nicaiu.dll" (ByVal taskHandle As Long) As Long
Declare Function DAQmxStopTask Lib "nicaiu.dll" (ByVal taskHandle As Long) As Long
Declare Function DAQmxClearTask Lib "nicaiu.dll" (ByVal taskHandle As Long) As Long

Declare Function DAQmxCreateDOChan Lib "nicaiu.dll" ( _
	ByVal taskHandle As Long, _
	ByVal lines As String, _
	ByVal nameToAssignToLines As String, _
	ByVal lineGrouping As Long _
) As Long

Declare Function DAQmxWriteDigitalLines Lib "nicaiu.dll" ( _
	ByVal taskHandle As Long, _
	ByVal numSampsPerChan As Long, _
	ByVal autoStart As Boolean, _
	ByVal timeout As Double, _
	ByVal dataLayout As Long, _
	ByRef writeArray() As Integer, _
	ByRef sampsPerChanWritten As Long, _
	ByRef reserved As Long _
) As Long

Declare Function DAQmxCreateAOVoltageChan Lib "nicaiu.dll" ( _
	ByVal taskHandle As Long, _
	ByVal physicalChannel As String, _
	ByVal nameToAssignToChannel As String, _
	ByVal minVal As Double, _
	ByVal maxVal As Double, _
	ByVal units As Long, _
	ByVal customScaleName As String _
) As Long

Declare Function DAQmxWriteAnalogF64 Lib "nicaiu.dll" ( _
	ByVal taskHandle As Long, _
	ByVal numSampsPerChan As Long, _
	ByVal autoStart As Boolean, _
	ByVal timeout As Double, _
	ByVal dataLayout As Boolean, _
	ByRef writeArray() As Double, _
	ByRef sampsPerChanWritten As Long, _
	ByRef reserved As Long _
) As Long

Declare Function DAQmxGetErrorString Lib "nicaiu.dll" ( _
	ByRef errorString As String, _
	ByVal bufferSize As Long _
) As Long

Declare Function DAQmxGetExtendedErrorInfo Lib "nicaiu.dll" ( _
	ByRef errorString As String, _
	ByVal bufferSize As Long _
) As Long

Declare Function DAQmxCfgSampClkTiming Lib "nicaiu.dll" ( _
	ByVal taskHandle As Long, _
	ByVal source As String, _
	ByVal sampleRate As Double, _
	ByVal activeEdge As Long, _
	ByVal sampleMode As Long, _
	ByVal sampsPerChanToAcquire As Long _
) As Long