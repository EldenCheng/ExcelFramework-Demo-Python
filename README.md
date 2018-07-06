# Elevate-Test-Performance
### Enviornment Setup
1. Install [JDK 8 or JDK 9](http://www.oracle.com/technetwork/java/javase/downloads/index.html) and set the System Environmental variable accordingly
2. Download [apache-jmeter-4.0 (zip or tar)](https://jmeter.apache.org/download_jmeter.cgi)
3. Unzip or ```tar -xvf``` to extract the downloaded file in a folder
### Test Plan on July release
1. ElevateScenarioTestPerformance.jmx is for Business cases load test, included below scenarios:
  * Invite multiple users
  * Multiple users sign up
  * Send a building message
  * Pull theme from BE
  * Delete users
2. ElevateTestPerformance.jmx is for API cases load test, included below APIs:
  * API Endpoints - auth
  * API Endpoints - base v1 - tenant
  * API Endpoints - base v1 - vendor
  * API Endpoints - base v1 - vendor (delete)
  * API Endpoints - base v1 - contact
  * API Endpoints - base v1 - property
  * API Endpoints - base v1 - resource
  * API Endpoints - base v1 - landlord
  * API Endpoints - base v1.1
  * API Endpoints - base v1.2
  * API Endpoints - booking
  * API Endpoints - event
  * API Endpoints - feed
  * API Endpoints - message
  * API Endpoints - permission
  * API Endpoints - property
  * API Endpoints - security
  * API Endpoints - static
  * API Endpoints - stats
### Setup Test Parameters
1. Set the testing target server and target VU number
  * Enter ```apache-jmeter-4.0/bin``` folder
  * Launch ```jmeter.bat``` or ```jmeter.sh``` to start Jmeter with GUI mode
  * Open the ```ElevateScenarioTestPerformance.jmx``` or ```ElevateTestPerformance.jmx``` in GUI mode
  * Click ```Testing Configuration Variables```
  * Set ```env``` to ```testing``` or ```staging``` in ```ElevateScenarioTestPerformance.jmx``` or ```ElevateTestPerformance.jmx```
  * Set target VU number in each threadgroup in ```ElevateScenarioTestPerformance.jmx```
  * Set ```numbersOfThreads``` to the target VU number in ```ElevateTestPerformance.jmx```
  * Saved the edited file(s)
2. Set other test parameters in config-testing.csv or config-staging.csv
  * url: The 1st parameter in the csv, it is for the target server's url
  * landlordUsername: The 2nd parameter in the csv, it is for the landlordAdmin user's name
  * landlordPassword: The 3rd parameter in the csv, it is for the landlordAdmin user's password
  * username: The 4th parameter in the csv, it is for the propertyAdmin user's name
  * password: The 5th parameter in the csv, it is for the propertyAdmin user's password
  * email: The 6th parameter in the csv, it is for the propertyAdmin user's email
  * emailUsername: The 7th parameter in the csv, it is for the new user's email fix part before @
  * emailDomain: The 8th parameter in the csv, it is for the new user's email part after @
  * propertyId: The 9th parameter in the csv, it need a valid propertyId which the propertyAdmin user can manger it
  * landlordId: The 10th parameter in the csv, it need a valid landlordId which the landlordAdmin user can manger it
  ### Test Run
  1. Launch the command-line window or terminal
  2. Enter ```apache-jmeter-4.0/bin``` folder
  3. Exectue ```jmeter -n -t [Test Plan file path] -l Results.jtl -e -r -j jmeter.log -o [Your Result folder]``` to let Jmeter run the load test in non-GUI mode
    * Please notice every time before you lanuch the load test, you need to make sure ```apache-jmeter-4.0/bin``` folder has no ```Result.jtl``` file and ```[Your Result folder]``` is an empty folder
    * The actually VU number is all the VU number you set in all threadgroup * all load generators if any
      For example, we plan to run ```ElevateTestPerformance.jmx``` and set ```numbersOfThreads``` to 4 with 2 load generators(one local, one remote), the totally VUs is 19(threadgroups) * 4(numberofThreads) * 2(numberofLoadgenerators) = 152
  ### Test Reslut
  1. Jmeter will auto generate a HTML report and place it in ```[Your Result folder]``` if the load test is start with parameter ```-o``` option
  2. Execute ```jmeter -g Results.jtl -o [Your Result folder]``` can generate the HTML reqport manually
  ### Trouble Shotting
  1. The HTTP Request responsed ```403 Forbidden```
    * The proertyAdmin user has no right to manage that property
    * The propertyAdmin user is a testing/staging env user, but used in staging/testing env
  2. The HTTP Request responsed ```500 Internal Server Error```
    * API server was really down
    * The request body parameter cause the API endpoint occurred unhandle error
  3. The HTTP Request responsed ```certificate_unkown...```
    * The temp Jmeter certificate expired, delete the old ```ApacheJMeterTemporaryRootCA.crt``` in certificate manager of the OS system, launch Jmeter in GUI mode, load a pre-set test plan ```Recording```, click ```HTTP(S) Test Script Recorder``` and click ```Start```, Jmeter will renew the certificate, then enter ```apache-jmeter-4.0/bin``` folder, install the new ```ApacheJMeterTemporaryRootCA.crt```
    * The API server not accept Jmeter's certificate, maybe can ask developer's help to shutdown the certificate checking on API server temporarily
    
