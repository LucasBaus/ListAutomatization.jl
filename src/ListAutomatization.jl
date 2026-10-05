module ListAutomatization
    using XLSX, DataFrames

    struct TempAlphaInfo
        letter::Vector{Char}
        names::Vector{Tuple{String, Int}}
        rowCount::Int
    end

    # This script take an entry XLSX file containing names and columns (no sort required)
    # It returns a new XLSX file, created based on the templates given (template.xlsx by default) sorting, categorizing etc...

    _alpha = collect('A':'Z')
    alpha = [[elt] for elt in _alpha[1:20]]
    push!(alpha, ['U', 'V'], ['W', 'X', 'Y', 'Z'])
    #min_alpha = lowercase.(alpha)
    initRow = 4
    initCol = 1
    maxRowCount = 61
    marginRow = 4

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

    function _main(entryFile::String ;template::String = "template.xlsx")
        template = "data/"*template
        df = DataFrame(XLSX.readtable(entryFile))
        for row in eachrow(df)
            row.nom = uppercase(row.nom)
        end
        sort!(df, :nom)

        L = TempAlphaInfo[]
        for letter in alpha
            sub_df = df[[elt[1] in letter for elt in df.nom], :]
            push!(L, TempAlphaInfo(letter, [(elt.nom, elt.colonne) for elt in eachrow(sub_df)], max(4, marginRow+size(sub_df, 1))))
        end

        name_exitFile = "$(split(entryFile, ".xlsx")[1])_format.xlsx"
        cp(template, name_exitFile, force=true)

        XLSX.openxlsx(name_exitFile, mode = "rw") do exit_file
            exit_file = exit_file[1]
            #TODO: Tricky part, this code is template exclusive (Use of Hard cell reference), change that automaticaly would be a pain in the ass
            row, col = initRow, initCol
            exit_file[getcell(1, 1)] = df[1, :Adresse]
            XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center")
            XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
            XLSX.setFont(exit_file, getcell(row, col); bold = true)
            for elt in L
                if row+elt.rowCount > maxRowCount
                    row = initRow
                    col += 2
                end
                #println(getcell(row, col))
                exit_file[getcell(row, col)] = changeLettersToString(elt.letter)
                XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center")
                XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
                XLSX.setFont(exit_file, getcell(row, col); bold = true)
                exit_file[getcell(row, col+1)] = exit_file[getcell(initRow, col+1)]
                XLSX.setAlignment(exit_file, getcell(row, col+1); horizontal="center")
                XLSX.setFill(exit_file, getcell(row, col+1); pattern = "solid", fgColor = "lightgrey")

                row += 1
                for name in elt.names
                    exit_file[getcell(row, col)] = name[1]
                    exit_file[getcell(row, col+1)] = name[2]
                    row += 1
                end
                row += elt.rowCount-length(elt.names)
            end
            # exit_file[getcell(row, col)] = "REBUTS"
            # XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center")
            # XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
            # XLSX.setFont(exit_file, getcell(row, col); bold = true)
        end
    end

    function julia_main()::Cint
        println("Test")
        println("Mes arguments : ", ARGS)
        return _main(ARGS[1])
        return 0
    end
end
