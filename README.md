This script generates a csv file with statistics on MS Teams spaces. 

For credentials, a file named uid.txt containing the clear text UID needs to be present in the directory. 
Additionally, a file named pass.txt containing the encrypted standard string of the account password. 

To create the encrypted standard string, use
read-host -assecurestring | converfrom-securestring | out-file -filepath .\pass.txt

