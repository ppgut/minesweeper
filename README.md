# minesweeper
excel minesweeper

Worksheet events used: SelectionChange, BeforeRightClick 

Required functionalisies couldn't be achived by direct use of those events. Instead, events are used to set the boolean switches bSelectionChanged and/or bRightClicked.
Then the switches are used to schedule required action 100 miliseconds after event was triggered.

Minesweeper requires quite challanging logic to be implemented to check the field which the user wants to reveal.
Recursive procedure 'UnHide' is used in this file to check selected filed as well as all the fields around it.

Conditions to be checked: [A] there is a mine on the field (game lost), [B] there is a number printed on the field (number of mines in direct surrounding), [C] field is empty (in this case surrounding fileds needs to be checked recursively to find the ones that are empty or contain number so they can be revealed as well)
