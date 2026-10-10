module ListAutomatization
    using XLSX, DataFrames, JSON3

    mutable struct Parameter
        cellHeight::Float64
        cellNameLength::Float64
        cellColLength::Float64
        cellMarginLength::Float64
        cellTrashLength::Float64
        withTrash::Bool

        marginRow::Int
    end

    function Parameter()
        return Parameter(
            11.75,
            22.,
            4.,
            2.,
            21.,
            true,
            1
        )
    end

    function null_Parameter()
        return Parameter(
            -1.,
            -1.,
            -1.,
            -1.,
            -1.,
            true,
            -1
        )
    end

    function copy_Parameter(p::Parameter)
        p_new= Parameter()
        for k in fieldnames(Parameter)
            setproperty!(p_new, k, getproperty(p, k)) 
        end
        return p_new
    end

    # Ensuite ici, une fonction modifiant un paramètre en fonction d'un JSON
    function json_update!(p::Parameter, json::String)
        json = JSON3.parse(json)
        for (k, v) in json
            setproperty!(p, k, v)
        end
    end

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

    p_base = Parameter()
    #min_alpha = lowercase.(alpha)
    initRow = 4
    initCol = 1

    const sheetWidth = 190
    const sheetHeight = 830

    #= Ici je veux réaliser un pruning des paramètres
        * Ce qui serait génial, c'est d'avoir un objet paramètre avec chaque élément initilaisé à "null"
        * Ensuite, si l'utilisateur fixe un paramètre, ion l'update
        * Chaque paramètre fixé est bloqué, pour les autres, on commence à la valeur par défaut, et on tente d'augmenter, jusqu'à dépassé la limite
        * Une fois, celà atteint, on connait les paramètre maximisant l'affichage sans laisser de colonnes vides.
        * Liste des paramètre à modifier dans l'ordre de préférence : 
            1. marginRow
            2. Si marginRow >= 5 (?)
                3. On redescend marginRow à 3
                4. cellHeight
                5. Puis marginRow de nouveau
    =#


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

    function checkEmptyCellLast(L::Vector{TempAlphaInfo}, maxRowCount::Int)
        s = maxRowCount
        col = 1
        for item in L
            if s - item.rowCount < 0
                col += 1
                s = maxRowCount
            else
                s -= item.rowCount
            end
        end
        return s - (col-2)*maxRowCount
    end

    function parameter_pruning(df::DataFrame, p_base::Parameter, p_utilisateur::Parameter, unfix_filednames; verbose::Bool = false)
        p_final = copy_Parameter(p_utilisateur)
        cpt = 0

        for item in unfix_filednames
            setproperty!(p_final, item, getproperty(p_base, item))
        end

        change = []
        stop = false

        while !stop
            cpt += 1
            println("Try number: $cpt")
            # Ici, calcul de si ca passe, si oui, on prune, si non, on stop, on va voir pour le retour en arrière si stop après :)

            maxRowCount = div(sheetHeight, p_final.cellHeight+0.25)
            L = TempAlphaInfo[]
            for letter in alpha
                sub_df = df[[elt[1] in letter for elt in df.nom], :]
                push!(L, TempAlphaInfo(letter, [(elt.nom, elt.colonne) for elt in eachrow(sub_df)], max(p_final.marginRow, p_final.marginRow+size(sub_df, 1))))
            end

            nb_empty = checkEmptyCellLast(L, Int(maxRowCount-3))
            println("Check empyt cell last: $nb_empty")
            println("Changes: $change")
            println("Max row count: $maxRowCount\n")
            if nb_empty < 0
                @assert length(change) > 0 "No possible configuration for the excel with the given parameter"
                if length(change)==2
                    stop = true
                elseif change[1][1] == :marginRow
                    p_final.marginRow = change[1][2]
                    change = []
                    push!(change, (:cellHeight, p_final.cellHeight))
                    p_final.cellHeight += 1
                else
                    stop = true
                end
            else
                # Ici, le changement de paramètre, la base du pruning
                change = []
                if (:marginRow in unfix_filednames) && p_final.marginRow >= 5
                    push!(change, (:marginRow, p_final.marginRow))
                    p_final.marginRow = p_base.marginRow
                    
                    if !(:cellHeight in unfix_filednames)
                        stop = true
                    else
                        push!(change, (:cellHeight, p_final.cellHeight))
                        p_final.cellHeight += 1
                    end
                elseif (:marginRow in unfix_filednames)
                    push!(change, (:marginRow, p_final.marginRow))
                    p_final.marginRow += 1
                else
                    if !(:cellHeight in unfix_filednames)
                        stop = true
                    else
                        push!(change, (:cellHeight, p_final.cellHeight))
                        p_final.cellHeight += 1
                    end
                end
            end
        end
        for (k, v) in change
            setproperty!(p_final, k, v)
        end
        return p_final
    end

    function _main(entryFile::String ; name_exitFile::String = "", parameters::String="{}", verbose::Bool = false)
        df = DataFrame(XLSX.readtable(entryFile))
        for row in eachrow(df)
            row.nom = uppercase(row.nom)
        end
        sort!(df, :nom)

        p_utilisateur = null_Parameter()
        json_update!(p_utilisateur, parameters)

        p_null = null_Parameter()

        unfix_filednames = Symbol[]
        for field in fieldnames(Parameter)
            if getproperty(p_utilisateur, field) == getproperty(p_null, field)
                push!(unfix_filednames, field)
            end
        end

        p_pruned = parameter_pruning(df, p_base, p_utilisateur, unfix_filednames)


        maxRowCount = Int(div(sheetHeight, p_pruned.cellHeight))
        verbose && println(sheetHeight, " - ", p_pruned.cellHeight)
        L = TempAlphaInfo[]
        for letter in alpha
            sub_df = df[[elt[1] in letter for elt in df.nom], :]
            push!(L, TempAlphaInfo(letter, [(elt.nom, elt.colonne) for elt in eachrow(sub_df)], max(p_pruned.marginRow, p_pruned.marginRow+size(sub_df, 1))))
        end

        if name_exitFile == ""
            name_exitFile = "$(split(entryFile, ".xlsx")[1])_format.xlsx"
        end
        verbose && println(name_exitFile)

        cp("data/TemplateNoMarge.xlsx", name_exitFile, force = true)

        XLSX.openxlsx(name_exitFile, mode = "rw") do exit_file
            exit_file = exit_file[1]

            for i = 1:maxRowCount
                for j = 1:8
                    verbose && println(i, j)
                    exit_file[getcell(i, j)] = ""
                end
            end

            XLSX.setRowHeight(exit_file, "A1:A$maxRowCount"; height=p_pruned.cellHeight)
            XLSX.setColumnWidth(exit_file, getcell(1, 1); width=p_pruned.cellNameLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 3); width=p_pruned.cellNameLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 5); width=p_pruned.cellNameLength)
            XLSX.setColumnWidth(exit_file, 8; width=p_pruned.cellTrashLength)
            XLSX.setColumnWidth(exit_file, 7; width=p_pruned.cellMarginLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 2); width=p_pruned.cellColLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 4); width=p_pruned.cellColLength)
            XLSX.setColumnWidth(exit_file, getcell(1, 6); width=p_pruned.cellColLength)

            #TODO: Tricky part, this code is template exclusive (Use of Hard cell reference), change that automaticaly would be a pain in the ass
            row, col = 4, 1
            verbose && println(row, col)
            XLSX.mergeCells(exit_file, "$(getcell(1, 1)):$(getcell(1+1, 1+7))")
            verbose && println(row, col)
            XLSX.setAlignment(exit_file, getcell(1, 1); horizontal="center", vertical = "center")
            verbose && println(row, col)
            exit_file[getcell(1, 1)] = df[1, :Adresse]
            verbose && println(row, col)
            XLSX.setFont(exit_file, getcell(1, 1); bold = true, color = "orange", size = 18)
            
            for elt in L
                if row+elt.rowCount > maxRowCount
                    row = initRow
                    col += 2
                end
                verbose && println(row, col)
                #verbose && println(getcell(row, col))
                verbose && println(row, col)
                exit_file[getcell(row, col)] = changeLettersToString(elt.letter)
                verbose && println(row, col)
                XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center", vertical = "center")
                verbose && println(row, col)
                XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
                verbose && println(row, col)
                XLSX.setFont(exit_file, getcell(row, col); bold = true)
                verbose && println(row, col)
                exit_file[getcell(row, col+1)] = "Col"
                verbose && println(row, col)
                XLSX.setAlignment(exit_file, getcell(row, col+1); horizontal="center", vertical = "center")
                verbose && println(row, col)
                XLSX.setFill(exit_file, getcell(row, col+1); pattern = "solid", fgColor = "lightgrey")

                row += 1
                for name in elt.names
                    verbose && println(row, col)
                    exit_file[getcell(row, col)] = name[1]
                    verbose && println(row, col)
                    XLSX.setAlignment(exit_file, getcell(row, col); vertical = "center")
                    verbose && println(row, col)
                    exit_file[getcell(row, col+1)] = name[2]
                    verbose && println(row, col)
                    XLSX.setAlignment(exit_file, getcell(row, col+1); horizontal="center", vertical = "center")
                    row += 1
                end
                row += elt.rowCount-length(elt.names)
            end
            

            XLSX.setBorder(exit_file, "A4:F$maxRowCount"; allsides = ["style" => "thin"])
            XLSX.setBorder(exit_file, "A1:H2"; allsides = ["style" => "thin"])
            XLSX.setBorder(exit_file, "H4:H$maxRowCount"; allsides = ["style" => "thin"])
            col, row = 8, 4
            verbose && println(row, col)
            exit_file[getcell(row, col)] = "REBUTS"
            verbose && println(row, col)
            XLSX.setAlignment(exit_file, getcell(row, col); horizontal="center", vertical = "center")
            verbose && println(row, col)
            XLSX.setFill(exit_file, getcell(row, col); pattern = "solid", fgColor = "orange")
            verbose && println(row, col)
            XLSX.setFont(exit_file, getcell(row, col); bold = true)
        end
        return name_exitFile
    end

    function main(args::Vector; json::String)
        name_exitFile = length(args) > 1 ? args[2] : ""
        return _main(args[1], name_exitFile = name_exitFile, parameters = json)
    end
end
