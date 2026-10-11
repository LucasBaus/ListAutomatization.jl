module ListAutomatization
    using XLSX, DataFrames, JSON3


    mutable struct ParameterValue{T}
        value::T
        enabled::Bool
    end

    struct Parameter
        cellHeight::ParameterValue{Float64}
        cellNameLength::ParameterValue{Float64}
        cellColLength::ParameterValue{Float64}
        cellMarginLength::ParameterValue{Float64}
        cellTrashLength::ParameterValue{Float64}
        withTrash::ParameterValue{Bool}

        marginRow::ParameterValue{Int}
    end

    const sheetWidth = 190
    const sheetHeight = 830

    function ParameterValue{T}(value::T) where T<: Any
        return ParameterValue{T}(value, false)
    end

    function Parameter(cellHeight::Float64, cellNameLength::Float64, cellColLength::Float64, cellMarginLength::Float64, cellTrashLength::Float64, withTrash::Bool, marginRow::Int)
        return Parameter(
            ParameterValue{Float64}(cellHeight),
            ParameterValue{Float64}(cellNameLength),
            ParameterValue{Float64}(cellColLength),
            ParameterValue{Float64}(cellMarginLength),
            ParameterValue{Float64}(cellTrashLength),
            ParameterValue{Bool}(withTrash),
            ParameterValue{Int}(marginRow),
        )
    end
    
    function Parameter()
        return Parameter(
            12.,
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
            getproperty(p_new, k).value = getproperty(p, k).value
            getproperty(p_new, k).enabled = getproperty(p, k).enabled
        end
        return p_new
    end

    # Ensuite ici, une fonction modifiant un paramètre en fonction d'un JSON
    function json_update!(p::Parameter, json::JSON3.Object)
        for (k, v) in json
            println(k)
            if k == :rowNameNumber
                println("In row name number")
                p.cellHeight.value = round(sheetHeight/(parse(Int, v)+3), digits = 2)
                p.cellHeight.enabled = true
                println(p.cellHeight)
            elseif k in fieldnames(Parameter)
                if getproperty(p, k).enabled
                    @warn "The value of the parameter $k is already fixed at : $(getproperty(p, k).value), this will be changed to $v"
                end
                getproperty(p, k).value = parse(typrof(getproperty(p, k).value), v)
                getproperty(p, k).enabled = true
            else
                @warn "The parameter $k is not defined and won't be applied"
            end
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

    function parameter_pruning(df::DataFrame, p_base::Parameter, p_utilisateur::Parameter; verbose::Bool = false)
        p_final = copy_Parameter(p_utilisateur)
        cpt = 0

        change = []
        stop = false

        while !stop
            cpt += 1
            println("Try number: $cpt")
            # Ici, calcul de si ca passe, si oui, on prune, si non, on stop, on va voir pour le retour en arrière si stop après :)

            maxRowCount = div(sheetHeight, p_final.cellHeight.value)
            L = TempAlphaInfo[]
            for letter in alpha
                sub_df = df[[elt[1] in letter for elt in df.nom], :]
                push!(L, TempAlphaInfo(letter, [(elt.nom, elt.colonne) for elt in eachrow(sub_df)], max(p_final.marginRow.value, p_final.marginRow.value+size(sub_df, 1))))
            end

            nb_empty = checkEmptyCellLast(L, Int(maxRowCount-3))
            println("Check empty cell last: $nb_empty")
            println("Changes: $change")
            println("Max row count: $maxRowCount\n")
            if nb_empty < 0
                @assert length(change) > 0 "No possible configuration for the excel with the given parameter"
                if length(change)==2
                    stop = true
                elseif (change[1][1] == :marginRow) && !(p_final.cellHeight.enabled)
                    p_final.marginRow.value = change[1][2]
                    change = []
                    push!(change, (:cellHeight, p_final.cellHeight.value))
                    p_final.cellHeight.value += 1
                else
                    stop = true
                end
            else
                # Ici, le changement de paramètre, la base du pruning
                # BUG !!!!!!!!!!!!!!!!!!!!!!!!!
                # Quand cellHeight enabled, il augmente la taille des cases quand même ! 
                change = []
                if !p_final.marginRow.enabled && p_final.marginRow.value >= 5
                    push!(change, (:marginRow, p_final.marginRow.value))
                    p_final.marginRow.value = p_base.marginRow.value
                    
                    if  p_final.cellHeight.enabled
                        stop = true
                    else
                        push!(change, (:cellHeight, p_final.cellHeight.value))
                        p_final.cellHeight.value += 1
                    end
                elseif !p_final.marginRow.enabled
                    push!(change, (:marginRow, p_final.marginRow.value))
                    p_final.marginRow.value += 1
                else
                    if p_final.cellHeight.enabled
                        stop = true
                    else
                        push!(change, (:cellHeight, p_final.cellHeight.value))
                        p_final.cellHeight.value += 1
                    end
                end
            end
        end
        for (k, v) in change
            getproperty(p_final, k).value = v
        end
        return p_final
    end

    function _main(entryFile::String ; name_exitFile::String = "", parameters::JSON3.Object = JSON3.Object(), verbose::Bool = false)
        df = DataFrame(XLSX.readtable(entryFile))
        for row in eachrow(df)
            row.nom = uppercase(row.nom)
        end
        sort!(df, :nom)

        p_utilisateur = Parameter()
        verbose && println("Avant Json update")
        json_update!(p_utilisateur, parameters)

        verbose && println("Avant Parameter pruning")
        p_pruned = parameter_pruning(df, p_base, p_utilisateur, verbose = verbose)


        maxRowCount = Int(div(sheetHeight, p_pruned.cellHeight.value))
        verbose && println(sheetHeight, " - ", p_pruned.cellHeight.value)
        L = TempAlphaInfo[]
        for letter in alpha
            sub_df = df[[elt[1] in letter for elt in df.nom], :]
            push!(L, TempAlphaInfo(letter, [(elt.nom, elt.colonne) for elt in eachrow(sub_df)], max(p_pruned.marginRow.value, p_pruned.marginRow.value+size(sub_df, 1))))
        end

        if name_exitFile == ""
            name_exitFile = "$(split(entryFile, ".xlsx")[1])_format.xlsx"
        end
        verbose && println(name_exitFile)

        if isfile(name_exitFile)
            rm(name_exitFile, force = true)
        end

        verbose && println("Avant cp")
        verbose && println(pwd())
        cp(joinpath(
            @__DIR__,
            "..",
            "data",
            "TemplateNoMarge.xlsx",
        ), name_exitFile)

        chmod(name_exitFile, 0o777)

        verbose && println("Avant création xlsx")
        verbose && println("FILE EXISTS: ", isfile(name_exitFile))
        verbose && println("FILE WRITABLE: ", iswritable(name_exitFile))
        verbose && println("DIR WRITABLE: ", iswritable(dirname(name_exitFile)))
        verbose && println("FILE READONLY: ", isreadonly(name_exitFile))
        verbose && println("PERMISSIONS: ", stat(name_exitFile), "\n", "PERMISSIONS: ", uperm(name_exitFile))
        XLSX.openxlsx(name_exitFile, mode = "rw") do exit_file
            exit_file = exit_file[1]

            for i = 1:maxRowCount
                for j = 1:8
                    verbose && println(i, j)
                    exit_file[getcell(i, j)] = ""
                end
            end

            XLSX.setRowHeight(exit_file, "A1:A$maxRowCount"; height=p_pruned.cellHeight.value-0.25)
            XLSX.setColumnWidth(exit_file, getcell(1, 1); width=p_pruned.cellNameLength.value)
            XLSX.setColumnWidth(exit_file, getcell(1, 3); width=p_pruned.cellNameLength.value)
            XLSX.setColumnWidth(exit_file, getcell(1, 5); width=p_pruned.cellNameLength.value)
            XLSX.setColumnWidth(exit_file, 8; width=p_pruned.cellTrashLength.value)
            XLSX.setColumnWidth(exit_file, 7; width=p_pruned.cellMarginLength.value)
            XLSX.setColumnWidth(exit_file, getcell(1, 2); width=p_pruned.cellColLength.value)
            XLSX.setColumnWidth(exit_file, getcell(1, 4); width=p_pruned.cellColLength.value)
            XLSX.setColumnWidth(exit_file, getcell(1, 6); width=p_pruned.cellColLength.value)

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
            verbose && println("Fin de la création du Excel")
        end
        verbose && println("Fin de _main")
        verbose && println("FILE CLOSED")
        return name_exitFile
    end

    function main(args::Vector; json::JSON3.Object = JSON3.Object(), verbose::Bool = false)
        name_exitFile = length(args) > 1 ? args[2] : ""
        return _main(args[1], name_exitFile = name_exitFile, parameters = json, verbose = verbose)
    end
end
