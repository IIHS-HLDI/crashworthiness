![](http://www.iihs.org/frontend/images/logohq.svg)
---

# Crashworthiness Processing Overview

IIHS crashworthiness tests are processed using the scripts that are contained within this repository. The directory structure of this repository must be maintained for these scripts to run.

- **Frontal Crash Tests**
    - The test profiles located in '_Frontal/Profiles_' are used to process moderate overlap, small overlap driver-side, and small overlap passenger side-tests.
    - Each test profile uses a different config file located in '_Frontal/ConfigFiles_' which contains the meta data needed to process each of the channels in the dummy and the vehicle.

- **Side and Head Restraint Tests**
    - Side crash and head restraint tests continue to use the more legacy processing scripts and have not been updated to use the newer test profile/config file model. Until side and rear crash programs are updated to the new processing model, the scripts in this repository are best used for reference to see how IIHS processes each test. Substantial modifications will be necessary in order to run the scripts.

## Modifications Needed to Run Scripts

- **Frontal Crash Tests**

    - For frontal tests, any node under the '_testSettings_' node that contains an attribute with a path should be modified for your use case. Channel names should also be modified to accommodate your naming convention.
    - IIHS uses a '_testTypeId_' to help select which test configuration file is used. Take a look at the '_ConfigFileFactory.VBS_' file to see how the ids are mapped to the config files. Change as needed. 
    - IIHS uses the following directory structure to store data for a test. If your case does not follow this structure you will have to modify the '_GetTestErrorLog_' method in the '_CrashTools.VBS_' file. 

            .
            +-- _DATA
            |   +-- _DAS
            |   +-- _DIAdem
            |   +-- _EDR
            |   +-- _EXCEL
            +-- _PHOTOS
            +-- _REPORTS
            +-- _VIDEO
            +-- info.txt

## Example: Processing the Hybrid III 50th Dummy

- Open DIAdem and load and save a data file. Change all channels that will be processed to waveform
- Using the H350 class:

        Call ScriptInclude(<path to H350Class.VBS>\H350Class.VBS")
        Dim DriverDummy : Set DriverDummy = (New H350)("11", iTestTypeId)

- Filtering the dummy:

        DriverDummy.FilterChannels

- Calculating injuries:

        DriverDummy.CalculateInjuryCriteria
