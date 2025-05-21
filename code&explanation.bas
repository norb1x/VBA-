' ==========================
' Borrow a Book
' ==========================

Option Compare Database
Option Explicit

Sub BorrowBook()
    Dim db As DAO.Database
    Dim rsBooks As DAO.Recordset
    Dim bookID As Long
    Dim borrowerName As String
    
    Set db = CurrentDb
    
    ' Ask for the book ID to borrow
    bookID = CLng(InputBox("Enter the ID of the book to borrow:", "Borrow Book"))
    If bookID = 0 Then Exit Sub
    
    ' Check book availability
    Set rsBooks = db.OpenRecordset("SELECT * FROM Books WHERE ID=" & bookID)
    
    If rsBooks.EOF Then
        MsgBox "Book with the specified ID does not exist.", vbExclamation
        rsBooks.Close
        Set rsBooks = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    If rsBooks!Status <> "Available" Then
        MsgBox "This book is currently borrowed.", vbInformation
        rsBooks.Close
        Set rsBooks = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    ' Ask for borrower's name
    borrowerName = InputBox("Enter the borrower's full name:", "Borrow Book")
    If borrowerName = "" Then Exit Sub
    
    ' Insert loan record
    db.Execute "INSERT INTO Loans (BookID, Borrower, LoanDate) VALUES (" & bookID & ", '" & borrowerName & "', Date())", dbFailOnError
    
    ' Update book status to "Borrowed"
    db.Execute "UPDATE Books SET Status = 'Borrowed' WHERE ID = " & bookID, dbFailOnError
    
    MsgBox "The book has been borrowed successfully.", vbInformation
    
    rsBooks.Close
    Set rsBooks = Nothing
    Set db = Nothing
End Sub


' ==========================
' Return a Book
' ==========================

Option Compare Database
Option Explicit

Sub ReturnBook()
    Dim db As DAO.Database
    Dim rsLoans As DAO.Recordset
    Dim loanID As Long
    
    Set db = CurrentDb
    
    ' Ask for the loan ID to return
    loanID = CLng(InputBox("Enter the Loan ID for return:", "Return Book"))
    If loanID = 0 Then Exit Sub
    
    Set rsLoans = db.OpenRecordset("SELECT * FROM Loans WHERE ID=" & loanID)
    
    If rsLoans.EOF Then
        MsgBox "Loan with the specified ID was not found.", vbExclamation
        rsLoans.Close
        Set rsLoans = Nothing
        Set db = Nothing
        Exit Sub
    End If
    
    ' Update book status to "Available"
    db.Execute "UPDATE Books SET Status = 'Available' WHERE ID = " & rsLoans!BookID, dbFailOnError
    
    ' Remove the loan record
    db.Execute "DELETE FROM Loans WHERE ID = " & loanID, dbFailOnError
    
    MsgBox "Book return completed successfully.", vbInformation
    
    rsLoans.Close
    Set rsLoans = Nothing
    Set db = Nothing
End Sub


' ==========================
' Search for Books
' ==========================

Option Compare Database
Option Explicit

Sub SearchBooks()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim searchTerm As String
    Dim sql As String
    
    Set db = CurrentDb
    
    searchTerm = InputBox("Enter part of the book title or author name:", "Search Books")
    If Trim(searchTerm) = "" Then Exit Sub
    
    sql = "SELECT ID, Title, Author, Status FROM Books " & _
          "WHERE Title LIKE '*" & searchTerm & "*' OR Author LIKE '*" & searchTerm & "*'"
    
    Set rs = db.OpenRecordset(sql)
    
    If rs.EOF Then
        MsgBox "No books matching the search criteria were found.", vbInformation
    Else
        Dim resultList As String
        resultList = "Found books:" & vbCrLf & vbCrLf
        
        Do While Not rs.EOF
            resultList = resultList & "ID: " & rs!ID & " | Title: " & rs!Title & _
                         " | Author: " & rs!Author & " | Status: " & rs!Status & vbCrLf
            rs.MoveNext
        Loop
        
        MsgBox resultList, vbInformation, "Search Results"
    End If
    
    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub
