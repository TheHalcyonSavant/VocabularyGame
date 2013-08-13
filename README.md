VocabularyGame - Better your elocution
===============================
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;This is a very simple test game that accumulates points based on your correctly guessed choices. The question is on the top of the window and you must choose only one from five random choices to proceed to the next question.  
Testing is the best approach to gain any knowledge and with this game it can't be easier.

##Installation
1. Download [dictionary.xlsm](dictionary.xlsm?raw=true) to some permanent location on your disk (like My Documents)
2. Download [VocabularyGame.msi](/VocabularyGame.msi?raw=true), install and run VocabularyGame from your desktop
3. When the open file dialog pops up, locate dictionary.xlsm
4. Enjoy, the game is started !

##Game Rules
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;The game is simply played by choosing only one option from the five *Radio Buttons*. The choice you made should correspond to the question (the light-green text on the top of the window). If you choose correctly, you gain **+5 points**. But, if you are wrong, you **loose all your points** that you have accumulated so far. In meanwhile, if the *Countdown timer* is on and its time passes, then you loose **-15 points** instead of all of them.
 
When you click the sound icon (next to the question) you can hear and learn how to pronounce the phrase. These sounds are downloaded into the *sounds/* folder.
 
When you gain **30 points** for the first time, then you've made the first record. To break it again you must pass the best record made so far. To see all records go to *File -> Records*. The records are stored inside *dat/{dictionaryWithNoExt}_records.dat* file (e.g. *dat/dictionary_records.dat*) and saved upon leaving the application.

##dictionary.xlsm
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;When you open the application for the first, an open dialog pops out and asks you to locate *dictionary.xlsm*. This is a Macro-Enabled Excel file that contains all the unique entries, important for this game to choose from. You must have installed **Microsoft Office 2007** or newer version to edit this file.  
This excel file contains 4 columns:  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **A** : **English** - this is the key column. The question is randomly generated from this column. Contains single word (e.g. *affix*), phrases (e.g. *bundle up*), also words with additional explanation in parenthesis like *maiden (adjective)*;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **B** : *Lexicon* - a meaningful explanations in English;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **C** : **Synonyms** - words with same or similar meanings;  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; **D** : Macedonian - direct Cyrillic translation.  
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Columns **B**, **C** and **D** can have more then one meaning separated by semicolon and new line (e.g. see *calf*). Also, those columns can be empty if a translation or explanation is not necessary. The 5 random answers are generated from these 3 columns. 

##Settings

####Answer Types
corresponds to 'dictionary.xlsm' columns: *Lexicon*, **Synonyms** and Macedonian. The 5 random choices are generated through limitation on these checked MenuItems (answer types). For example, if you check Lexicon and Synonyms, but uncheck Macedonian, then you will NOT see Macedonian (Cyrillic) words in the 5 random choices from the next question.
>If some random choice is Lexicon, then it's displayed in *Italic*. And if it's Synonym it is displayed in **Bold**.

####Auto-Pronounce question
Automatically pronounce the question on every new round.

####Don't show me choices I guessed more then
This setting limits the repetition of displayed question-answer pairs. For example if you choose "Don't show me choices I guessed more then = 3 times" and if you have guessed question "virtue" with answer "merit" twice already, then this question with this particular answer will not be displayed again.  
The repetitions history of these pairs are saved inside *dat/{dictionaryWithNoExt}_repeats.dat* file (e.g. *dat/dictionary_repeats.dat*).  
This setting can be very handy if you want to filter very known words, but you want to avoid removing them from the excel file.

####Language
User Interface localization language. Needs restart for this setting to take effect.

####Reset All Settings
Reset all settings in the main menu "Settings" to its default values. The default values can be edited inside *VocabularyGame.exe.config* file under *configuration > userSettings*

Thanx to
------------
[Pavel Chuchuva](http://stackoverflow.com/a/1134340) for GifImage.cs  
[Google Dictionary](http://en.wikipedia.org/wiki/Google_Dictionary) for making this game "sounds" better
######Other repositories:
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Use [ExcelApp]() library if you want to compile this code successfully
