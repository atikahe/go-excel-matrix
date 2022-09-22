# goexcelmatrix

Project goexcelmatrix generates matrix into excel and enables writing value into a cell given the row and column name.

Early development stage, hence somewhat obscure code.

## Installation
```sh
go get github.com/atikahe/go-excel-matrix
```

## Usage
```go
import github.com/atikahe/go-excel-matrix

package main() {
    // Create new file
    excel := goexcelmatrix.NewFile("row-", "col-")

    // Initialize rows and columns
    rows := []string{"Student-1", "Student-2", "Student-3"}
    cols := []string{"Math", "Science", "History"}

    excel.SetRows(rows).SetColumns(cols)

    // Input data
    scores := [][]string{
        {"Student-1", "Math", "A"},
        {"Student-2", "Social Science", "B"},
        {"Student-3", "History", "B"},
    }

    for _, score := range score {
        excel.SetValue(score)
    }

    // Save file
    excel.Save("AP_Score_2022")
}
```

## Testing
Testing will go here.

## Contribution
Contributions will go here.

## Project Status
Early development.

