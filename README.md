\# Excel Arcade



A collection of fully playable games built entirely inside Microsoft Excel using VBA.



The project currently includes:



\- Chess

\- Blackjack



No external game engine or graphics library is required. The playable games are contained in the Excel workbook, while the exported VBA source code is included separately so the implementation can be reviewed directly on GitHub.



\## Screenshots



\### Chess



!\[Chess](screenshots/Chess.png)



\### Blackjack



!\[Blackjack](screenshots/Blackjack.png)



\## Chess



A fully playable chess implementation running entirely inside an Excel worksheet.



\### Features



\- Legal movement for all chess pieces

\- Interactive legal move highlighting

\- Turn-based gameplay

\- Captures

\- King safety validation

\- Check detection

\- Checkmate detection

\- Stalemate detection

\- Pinned-piece handling

\- Discovered checks

\- Double-check handling through legal move validation

\- Kingside castling

\- Queenside castling

\- Castling-right tracking after king or rook movement

\- En passant

\- Pawn promotion

\- Underpromotion to rook, bishop, or knight

\- Threefold repetition draw claims

\- Fivefold repetition automatic draw

\- 50-move rule draw claims

\- 75-move rule automatic draw

\- Standard Algebraic Notation (SAN) move history

\- Last-move highlighting

\- Undo

\- New Game



\## Blackjack



A multiplayer Blackjack implementation using a persistent two-deck shoe.



Unlike a simple random-card generator, the game creates and shuffles two complete physical decks in memory. Cards are then drawn from that finite 104-card shoe until a new shoe is created.



\### Features



\- Real 104-card shoe

\- Two complete 52-card decks

\- Fisher-Yates shuffle

\- Persistent shoe across multiple rounds

\- Changing probabilities as cards leave the shoe

\- New Shoe control

\- Up to four players

\- Individual IN / OUT player selection

\- Optional dealer mode

\- Dealer hole card is drawn immediately but remains hidden

\- Hit

\- Stand

\- Double

\- Split

\- Split aces

\- Soft ace calculation

\- Natural blackjack detection

\- Split 21 treated separately from natural blackjack

\- Dealer stands on soft 17

\- Automatic dealer play

\- Win, loss, push, blackjack, and bust evaluation

\- Practice mode when the dealer is OUT



Because the game uses a finite two-deck shoe, card availability is preserved correctly across rounds.



For example, two decks contain exactly eight aces. If all eight aces have already been dealt, another ace cannot appear until the player creates a new shoe.



\## Project Structure



```text

excel-arcade-vba/

|

|-- ExcelArcade.xlsm

|-- README.md

|

|-- src/

|   |-- ChessEngine.bas

|   `-- BlackjackEngine.bas

|

|-- ChessBoard.cls

|-- Blackjack.cls

|-- ThisWorkbook.cls

|

`-- screenshots/

&#x20;   |-- Chess.png

&#x20;   `-- Blackjack.png

```



\## Running the Project



1\. Download `ExcelArcade.xlsm`.

2\. Open the file using Microsoft Excel for Windows.

3\. Enable macros when prompted.

4\. Open either the `Chess Board` or `Blackjack` worksheet.

5\. Play directly through the worksheet interface.



The project requires the desktop version of Microsoft Excel with VBA support.



\## Source Code



The playable workbook is provided as:



`ExcelArcade.xlsm`



The VBA source code is also exported separately for easier inspection on GitHub.



\### Core Modules



\- `src/ChessEngine.bas`

\- `src/BlackjackEngine.bas`



\### Worksheet and Workbook Event Modules



\- `ChessBoard.cls`

\- `Blackjack.cls`

\- `ThisWorkbook.cls`



\## Technical Highlights



The project demonstrates:



\- Event-driven programming with VBA

\- Game-state management

\- Rule-based validation

\- Interactive worksheet interfaces

\- State-dependent user controls

\- Algorithmic move generation

\- King-safety simulation

\- Finite-deck simulation

\- Fisher-Yates shuffling

\- Persistent state across multiple rounds

\- Unicode-based game rendering

\- Excel cell formatting as a lightweight user interface

\- Multi-game architecture inside a single workbook



\## Why Excel?



The project explores how far Microsoft Excel can be pushed beyond traditional spreadsheet use.



Both games use Excel cells as the interface while VBA handles the underlying game logic, state management, validation, and interaction.



No external game engine is used.



\## Built With



\- Microsoft Excel

\- VBA

