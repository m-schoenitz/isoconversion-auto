function parameters =loadParameters(conf_file, sheet, headTopLeft, dataTopLeft)
    exl = actxserver('excel.application'); % Create a COM (Excel Component Object Model) server
    [exlWkbk exlFile exlSheet] = getSheets(exl, conf_file, sheet);
    [parameterHeads num]= getParameterHeads(exlSheet,headTopLeft);
    parameters=loadParams(parameterHeads, num, exlSheet, dataTopLeft);
    exlWkbk.Close % Close workbooks object, descending objects will be terminate automatically
    exl.Quit % Close COM server

    function [exlWkbk exlFile exlSheet] = getSheets(exl, conf_file, sheet)
        % After creating this server, it will be able to manipulate Excel objects as in the Excel VB scripts
        exlWkbk = exl.Workbooks; % Create a workbooks object using excel COM (Similar to )
        % exlFile = exlWkbk.Open('C:\Documents and Settings\KC\Desktop\Study Aboard\NJIT works\Master Project\Isoconversion Coding\Simulated-dsc-data\simulated-01Kmin.csv');
        exlFile = exlWkbk.Open(conf_file); % Manipulate workbooks object, using "open" method to open an existing workbook
        exlSheet = exlFile.Sheets.Item(sheet); % Create Excel Sheet object to read specific sheets.
    end
    function [parameterHeads num]= getParameterHeads(exlSheet, HeadTopLeft)
        %HeadTopLeft is in A1 to Z9
        headColName =HeadTopLeft(1);
        
%         eobj = exlSheet.Columns.End(1); % Grab sheet information
%         eobj2 = exlSheet.Columns.End(2); % Grab sheet information
%         eobj3 = exlSheet.Columns.End(3); % Grab sheet information
%         eobj4 = exlSheet.Columns.End(4); % Grab sheet information
%         numrows = eobj.row; % Determine how many rows are there in the sheet
%         dat_range = [HeadTopLeft ':' headColName num2str(numrows)]; % Define data range to read
%         rng = exlSheet.Range(dat_range); % get the head of parameters
        rng = exlSheet.Range(headTopLeft); % get the head of parameters
        %num = numrows-1;
        %num = 29;
        parameterHeads = rng.Value; % Read the head of parameters
        num=size(parameterHeads,1);
    end

    function parameters=loadParams(parameterHeads, num, exlSheet, dataTopLeft)
        %dataTopLeft is in A1 to ZZ9
        s = length(dataTopLeft);
        startRow = int16(str2num(dataTopLeft(s))) ;  %the last char is row num
        dataCol = dataTopLeft(1:s-1);
        n=1;
        while n < num
            head = parameterHeads{n};
            inc = 1; % defaut will go to the next line
            switch head
                case 'inputDataParameter'
                    [parameters.inputDataParameters inc] = getInputFiles(exlSheet, dataCol, startRow + n );
                case 'outputParameter'
                    [parameters.outputExlParams inc] = getOutputFile(exlSheet, dataCol, startRow + n );
                case 'progressParameters'
                    [parameters.progressParameters inc] = getProgressParameters(exlSheet, dataCol, startRow + n );
                case 'preProcessingParameters'
                    [parameters.preProcessingParameters inc] = getPreProcessingParameters(exlSheet, parameterHeads, n+1, dataCol, startRow+n);
            end
            n = n + inc;
        end
    end
    function [inputDataParameters inc] = getInputFiles(exlSheet, dataCol, startRow )
        numCell = [dataCol int2str(startRow)];
        num = int16(exlSheet.Range(numCell).Value);
        dirCell = [dataCol int2str(startRow+1)];
        dir = exlSheet.Range(dirCell).Value;
        fileRange = [dataCol int2str(startRow+2) ':' dataCol int2str(startRow + 1 + num)];
        files = exlSheet.Range(fileRange).Value;
        inc = 3 + num;  %skip 3+num line
        inputDataParameters.num = num;
        inputDataParameters.dir = dir;
        inputDataParameters.files = files;
    end 
    function [outputExlParams inc] = getOutputFile(exlSheet, dataCol, startRow )
        fileCell = [dataCol int2str(startRow)];
        file = exlSheet.Range(fileCell).Value;
        numCell = [dataCol int2str(startRow+1)];
        num = int16(exlSheet.Range(numCell).Value);
        sheetLocRange = [dataCol int2str(startRow+2) ':' getFollowColName(dataCol, 1) int2str(startRow + 1 + num)];
        sheetLocs = exlSheet.Range(sheetLocRange).Value;
        inc = 3 + num; %skip 3+num line
        outputExlParams.num = num;
        outputExlParams.file = file;
        outputExlParams.sheetLocation = sheetLocs;
    end 
    function [progressParameters inc] = getProgressParameters(exlSheet, dataCol, startRow )
        Range = [dataCol int2str(startRow) ':' dataCol int2str(startRow + 2 )];
        Prog = exlSheet.Range(Range).Value;
        progressParameters.startAlpha = Prog{1};
        progressParameters.delta = Prog{3};
        progressParameters.endAlpha = Prog{2};
        inc=4; %skip 4 line
    end
    function [preProcessingParameters inc] = getPreProcessingParameters(exlSheet, parameterHeads, headInd, dataCol, startRow )
        numCell = [dataCol int2str(startRow)];
        num = int16(exlSheet.Range(numCell).Value);
        pcell{num}=0;
        for i=1:num
            pcell{i}=getOnePreProcessingParameters(exlSheet, parameterHeads{headInd+i}, dataCol, startRow+i);
        end
        preProcessingParameters = pcell;
        inc=2 + num; %skip 2+num line
    end
    function processingParameters  = getOnePreProcessingParameters( exlSheet, parameterHead, dataCol, startRow)
        processingParameters.name = parameterHead;
        switch parameterHead
            case 'baselineCorrection'
                processingParameters.criteria = getRowArray(exlSheet,dataCol, startRow,3);
            case 'tempRange'
                processingParameters.boundary = getRowArray(exlSheet,dataCol, startRow,2);
            case 'smooth'
                numCell = [dataCol int2str(startRow)];
                processingParameters.num = int16(exlSheet.Range(numCell).Value);
            case 'tg2progress'
                numCell = [dataCol int2str(startRow)];
                processingParameters.maxInc = exlSheet.Range(numCell).Value;
            case 'dsc2progress'
        end
    end
    function bounds = getRowArray(exlSheet,dataCol, startRow, num)
        Range = [dataCol int2str(startRow) ':' getFollowColName(dataCol, num -1) int2str(startRow  )];
        bb = exlSheet.Range(Range).Value;
        bounds = cell2mat(bb); % Transfer value format
    end

    function Range = getRange(startCol, startRow, col, row)
        Range = [startCol int2str(startRow) ':' getFolowColName(startCol, col -1) int2str(startRow  +row -1)];
    end
    function ColName =getColName(colNum)
        %colNum should be a interge rather than float or double
        H='ABCDEFGHIJKLMNOPQRSTUVWXYZ';
        if colNum <27
            ColName = H(colNum);
        else
            a = colNum/26;
            b = colNum - a *26;
            ColName = [H(a) H(b)];
        end
    end
    function ColNum =getColNum(colName)
        Name =upper(colName);
        ColNum = Name(1) - 'A' +1;
        if size(colName) >1
            ColNum = ColNum*26 + Name(2) - 'A' +1;
        end
    end
    function followColName = getFollowColName(colName, n)
        m = getColNum(colName);
        followColName = getColName(m +n);
    end
end