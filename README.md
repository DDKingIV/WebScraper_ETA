# WebScraper_ETA

This python script(the one to use is ETA_Updater_Undetected.py, which takes steps to prevent being detected) takes an excel file with a list of all shipment for which we want to know the estimated time of arrival from the shipping lines websites.
Then it repeatedly scrapes the web to get the date from the shipping line searching with the container number.
It then outputs a file containing the Estimated arrival date for each shipment

YOU NEED A CHROMEDRIVER TO RUN THIS AND YOU NEED TO UPDATE THE LOCATION OF IT IN THE "Variables.py" File, where you will find also all the references to the HTML elements used by the script.
