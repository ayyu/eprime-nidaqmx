# E-Prime NI-DAQmx

E-Basic code for exposing and using National Instruments NI-DAQmx functions in an E-Prime script.

Luckily, E-Basic is a stripped down version of Visual Basic 6.0 so it's not hard to declare functions available in the DLL for use in experiments.

## Usage

Requires a copy of the `nicaiu.dll`, from the [NI DAQmx driver](https://www.ni.com/en-il/support/downloads/drivers/download.ni-daqmx.html). Copy and paste the DLL into the same directory as your `.es3` file.

See `UserScript.vb` for an example of what to include in your E-Prime UserScript in order to write to the NI device.

See `SampleFunctions.vb` for an example of some wrapper functions I've written in the past to make it easier to work with the declared DLL functions. You would include this in your User Script underneath the contents of the `UserScript.vb` file.

I haven't included any examples of reading from the NI device but it shouldn't be very hard.

## See also

- [NI DAQmx Constants](https://nidaqmx-python.readthedocs.io/en/latest/constants.html) for values of other constants in the library should you need them (I didn't declare all of them in my UserScript)
- [NI-DAQmx C Reference Help](https://zone.ni.com/reference/en-XX/help/370471AM-01/) for all available functions in the DLL
