This .DLL is used by WinGPIB, see WinGPIBproj.

NOTE: Installation .MSI for WinGPIB is located in one of the other Repositories, here: \WinGPIBadditionalinfo\WinGPIB-AnyCPU_All-SetupFiles

Here is a list of the current interface that are confirmed WinGPIB compatible:
GPIBDevice_NINET 	     - NI boards
GPIBDevice_ADLink 	   - ADLink boards
GPIBDevice_gpib488 	   - Keithley, MCC and older NI boards
VisaDevice 		         - Generic interface (GPIB,USB etc.) via Visa
COM Port 		           - Generic serial ports
Prologix		           - Prologix USB and Ethernet
NI-GPIB-232CT-A	       - Serial
Kofen's PoE Ethernet   - Prologix compatible
XYPHRO UsbGpib         - Visa

Here is a link to Pawel's original IODevices.dll code that I had used as a basis for my mods:
https://github.com/pawel-wzietek/IODevices
