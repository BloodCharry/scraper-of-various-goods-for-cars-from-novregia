a script that parses goods for cars from Norwegian websites and translates into English, enters data into Excel tables, the Python programming language and the following libraries were used: Beautiful Soup, Requests, Selenium, OpenPyXL and Googletrans.

Description of the script operation:

In the script, there are protections against fool users who like to write different nonsense instead of the necessary data).

The script asks how many categories need to be parsed on this resource, informing about the maximum number of categories.

Using the Requests library to send GET requests to these sites and get HTML pages with information about products for cars, selenium is used for some sites.

Using the Beautiful Soup library to parse HTML pages and extract the necessary information, for example, names, prices, product descriptions, etc.

Using the Googletrans library to translate the received data into the selected language (for example, Russian).

Using the OpenPyXL library to create and populate Excel tables with the received data.

Saving the created Excel tables to the selected directory.

If necessary, you can use additional libraries to automatically clear data from unwanted characters or to optimize the speed of the script.

At the end of the script, a message is displayed that the process has been completed successfully and the data tables are saved in the specified directory.

I do not debug the code in open access, this is my job, sorry).
