# VBA Library Management Tool – Technical Report

## Functional Summary
The system manages book loans through a set of well-separated VBA subroutines:
- `ResetWypozyczen`: Clears all entries from the `Loans` table and marks all books as "Dostępna".
- `MainProcess`: Central user interface, provides options for sync, loan, return, or book search.
- `ListAllTables`: Prints all table names from the current Access DB to the Immediate Window.

## Modularity
Each logical function is encapsulated in its own subroutine, improving code maintainability and making the system adaptable to new features.

## Use Case
Designed for small libraries or educational environments needing a basic but functional lending management interface inside Access.
