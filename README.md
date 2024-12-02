Excel to PGN generator
----------------------

This program accepts an excel file, as found in chess-results.com, and produces a PGN file suitable for uploading to lichess as a study. It is useful when PGN files are not available and needs to be entered by hand. It works with the excel file, found on the page, reached after clicking on a player's name. It saves a bit of time and perhaps prevents errors.

It is not as robust as it could be. In particular, I did not try to handle the case where the player of interest had a "bye" round, or won a game by fortfeit. I'll burn those bridges when I come to them.

It assumes that the names are in Turkish. If it is more appropriate to assume that they are in English, use "-e" command line switch. It also assumes that you downloaded from the Turkish version of the page.

The code "generate-pgn.py" is written by chatGPT. It came with a few bugs that I fixed with more prompting, but it was still faster than writing on my own.
