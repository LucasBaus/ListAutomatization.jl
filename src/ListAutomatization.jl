module ListAutomatization
    using XLSX, DataFrames, JuliaC

    struct TempAlphaInfo
        letter::Vector{Char}
        names::Vector{Tuple{String, Int}}
        rowCount::Int
    end

    # This script take an entry XLSX file containing names and columns (no sort required)
    # It returns a new XLSX file, created based on the templates given (template.xlsx by default) sorting, categorizing etc...

    _alpha = collect('A':'Z')
    alpha = [[elt] for elt in _alpha[1:8]]
    push!(alpha, ['I', 'J'])
    append!(alpha, [[elt] for elt in _alpha[11:20]])
    push!(alpha, ['U', 'V'], ['W', 'X', 'Y', 'Z'])
    #min_alpha = lowercase.(alpha)
    initRow = 4
    initCol = 1
    maxRowCount = 69
    marginRow = 1

    cellHeight = 12
    cellNameLength = 21
    cellColLength = 4
    cellMarginLength = 2
    cellTrashLength = 13
    withTrash = true
    
    function getcell(row::Int, col::Int)
        return "$(_alpha[col])$row"
    end

    function changeLettersToString(letters::Vector{Char})
        s = "$(letters[1])"
        for elt in letters[2:end]
            s *= " - $elt"
        end
        return s
    end

    function _main(entryFile::String)
        df = DataFrame(XLSX.readtable(entryFile))
        for row in eachrow(df)
            row.nom = uppercase(row.nom)
        end
        sort!(df, :nom)

        L = TempAlphaInfo[]
        for letter in alpha
            sub_df = df[[elt[1] in letter for elt in df.nom], :]
            push!(L, TempAlphaInfo(letter, [(elt.nom, elt.colonne) for elt in eachrow(sub_df)], max(2, marginRow+size(sub_df, 1))))
        end

        name_exitFile = "$(split(entryFile, ".xlsx")[1])_format.xlsx"

        XLSX.openxlsx(name_exitFile, mode = "w") do exit_file
            exit_file = exit_file[1]

            for i = 1:maxRowCount
                for j = 1:8
                    exit_file[getcell(i, j)] = ""
                end
            end

            XLSX.setRowHeight(exit_file, "A1:A$maxRowCount"; height=cellHeight)
            XLSX.setColumnWidth(exit_file, getcell(1, 1); width=cellNameLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 3); width=cellNameLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 5); width=cellNameLength)
            XLSX.setColumnWidth(exit_file, 8; width=cellTrashLength)
            XLSX.setColumnWidth(exit_file, 7; width=cellMarginLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 2); width=cellColLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 4); width=cellColLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 6); width=cellColLength)

            #TODO: Tricky part, this code is template exclusive (Use of Hard cell reference), change that automaticaly would be a pain in the ass
            row, col = 4, 1
            XLSX.mergeCells(exit_file, "$(getcell(1, 1)):$(getcell(1+1, 1+7))")
            XLSX.setAlignment(exit_file, getcell(1, 1); horizontal="center", vertical = "center")
            exit_file[getcell(1, 1)] = df[1, :Adresse]
            XLSX.setFont(exit_file, getcell(1, 1); bold = true, color = "orange", size = 18)
            
            for elt in L
                if row+elt.rowCount > maxRowCount
                    row = initRow
                    col += 2
                end
                #println(getcell(row, col))
                exit_file[getcell(row, col)] = changeLettersToString(elt.letter)
                XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center", vertical = "center")
                XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
                XLSX.setFont(exit_file, getcell(row, col); bold = true)
                exit_file[getcell(row, col+1)] = "Col"
                XLSX.setAlignment(exit_file, getcell(row, col+1); horizontal="center", vertical = "center")
                XLSX.setFill(exit_file, getcell(row, col+1); pattern = "solid", fgColor = "lightgrey")

                row += 1
                for name in elt.names
                    exit_file[getcell(row, col)] = name[1]
                    exit_file[getcell(row, col+1)] = name[2]
                    row += 1
                end
                row += elt.rowCount-length(elt.names)
            end
            

            XLSX.setBorder(exit_file, "A4:F$maxRowCount"; allsides = ["style" => "thin"])
            XLSX.setBorder(exit_file, "A1"; allsides = ["style" => "thin"])
            XLSX.setBorder(exit_file, "H4:H$maxRowCount"; allsides = ["style" => "thin"])
            # exit_file[getcell(row, col)] = "REBUTS"
            # XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center", vertical = "center")
            # XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
            # XLSX.setFont(exit_file, getcell(row, col); bold = true)
        end
    end
    
    function main(args::Vector{String})
        return _main(args[1])
    end

end
