Excel to PGN generator
----------------------

This program accepts an excel file, as found in chess-results.com, and produces a PGN file suitable for uploading to lichess as a study. It is useful when PGN files are not available and needs to be entered by hand. It works with the excel file, found on the page, reached after clicking on a player's name. It saves a bit of time and perhaps prevents errors.

It is not as robust as it could be. In particular, I did not try to handle the case where the player of interest had a "bye" round, or won a game by fortfeit. I'll burn those bridges when I come to them.

It assumes that the names are in Turkish. If it is more appropriate to assume that they are in English, use "-e" command line switch. It also assumes that you downloaded from the Turkish version of the page. If you find this useful and want

"The code 'generate-pgn.py' is written by chatGPT. It came with a few bugs that I fixed with more prompting, but it was still faster than writing on my own." This comment was correct for the initial version, but since then the code underwent further significant changes done by hand. Still, I think it is fair to give credit, where credit is due.

If you find this useful and want to make changes (more languages with proper internationalization, make it more robust to handle edge cases etc.), feel free to send a pull request. I use this code regularly and I welcome all improvements.

Typical Usage
=============

1. Go to the tournament page, in chess-results.com, e.g., visit https://chess-results.com/Tnr1185552.aspx?lan=8
Note that "lan=8" (that is choosing Turkish) is necessary for the program to function properly.
1b. If the tournament is old, you may need to click a few more links to get the participants list.
2. Choose a person from the list, e.g., I am number 22 on that list, clicking my name will take you to https://chess-results.com/tnr1185552.aspx?lan=8&art=9&fed=TUR&snr=22
3. In that page, choose to download the excel file, again for my case this would be following the link https://chess-results.com/tnr1185552.aspx?lan=8&zeilen=0&art=9&fed=TUR&snr=22&prt=4&excel=2010
4. Let's say that the excel file is saved under the name ato-aksaray-mayis-2025.xlsx You then run the program with the command
python gen-pgn2.py  -f ato-aksaray-mayis-2025.xlsx -o ato-aksaray-mayis-2025.pgn
This should generate a PGN file, which you can then edit to add moves.
4b. It is a valid PGN file, so you can directly upload it to a Lichess study, and add moves there. However, note that as of this writing, you cannot copy paste moves into a lichess study; if you have the moves somewhere else, you need to put them in the PGN file, and then upload.

