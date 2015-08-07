About the Reversi Game:

Reversi is a game played between human and computer on a board of size 10 x 10. The user (yellow colour) always plays first and then computer plays with red colour. 

	The player plays the game by placing yellow squares on empty cells of the board. Yellow square can be placed at an empty position if there are only red squares (at least one, no empty spaces) between it and any other yellow square. A chance is passed on to other player if there is no possibility of any placement.

	The crux of the game lies in flanking opponent's squares. Flanking is achieved by entrapping squares of the opponent's colour in between the currently placed square and a previously placed self-owned square. All these squares must be in a straight line, vertically, horizontally or diagonally in any orientation. The opponent's squares so trapped, change colour to the player's own colour. 

Flanking is not cascaded. That is, when a red square is flanked and changes colour to yellow, it does not trigger the flanking of other red squares. A chain of mixed squares or gaps is not considered as a flank. 

To be precise, a flank is defined as a sequence of squares having following properties: 

1. All the squares in the flank must be in a straight line (vertical, horizontal, or diagonal). 
2. There must be at least three squares in the flank. 
3. The flank must not have any gaps in it. 
4. The ends of the flank must be of the same colour, say C. 
5. All the squares in the flank, except its end- squares, must be of a different colour than C. 
6.  One of the ends of the flank must be the square placed in the current move.

At the end of the game (i.e. when all the squares of the board are filled), the player who owns more squares wins the game. It's a draw if there is a tie between the number of squares of player and the computer.

