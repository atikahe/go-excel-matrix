# goexcelmatrix

Project goexcelmatrix generates matrix into excel and enables writing value into a cell given the row and column name.

Early development stage, hence being somewhat obscure.

## Installation
```sh
go get github.com/atikahe/go-excel-matrix
```

## Usage
```go
package main

import "github.com/atikahe/go-excel-matrix"

package main() {
    // Create new file
    excel := goexcelmatrix.NewFile()

    // Initialize rows and columns
    rows := []string{"Student-1", "Student-2", "Student-3"}
    cols := []string{"Math", "Social Science", "History"}

    excel.SetRows(rows).SetColumns(cols)

    // Input data
    scores := [][]string{
        {"Student-1", "Math", "A"},
        {"Student-2", "Social Science", "B"},
        {"Student-3", "History", "B"},
    }

    for _, score := range scores {
        excel.SetValue(score)
    }

    // Save file
    excel.Save("AP_Score_2022")
}
```

## To Do
- Param validation & error handling
- Automate test

## Testing
Testing will go here.

## Contribution
Contributions will go here.

## Project Status
Early development.

