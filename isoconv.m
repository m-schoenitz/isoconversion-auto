% Basic steps of running the program 
% 1.	using MatlabR2007a or above

% 2.	filling up the configure file (config_TGA.xls) for all necessary
% information of running the program (check the comment of each cell for
% the meaning) 

% 3.	if raw data has big noise, data smoothing probably will be
% needed before the program is run. 

% 4.	opening Matlab, typing the following command to run the program:
% main 'config_TGA.xls' 'sheet1' 'A2:A28' 'B2' 

% 5.	generated figures will be the same sequence as input files,
% respectively.  

% 6.	all calculated results can be found from the output file of current
% directory.  

function Eava = isoconv(conf_file, sheet, range, dataTopLeft)

pause on;

% the main function serves as Ea calculation

	conf_file = [pwd '/' conf_file];

    params = loadParameters(conf_file, sheet, range, dataTopLeft);
	
	% this line looks strange:
    params.boundOfEa = [1000 1000000 1000];  % boundary of Ea is 100-10M J/mol

	params.method = 'Vyazovkin2000';  % hardcode the method name now
    [rawData, mass]= dataLoading(params.inputDataParameters);
    numOfTest = size(rawData,1);	% number of independent experiments
    expectedAlpha = createExpectedProgressVector(params.progressParameters);
    numOfRes = size(expectedAlpha,2);	% number of values of alpha where
										% Ea will be calculated
										% = "number of results"
										
    rangeData{numOfRes,numOfTest}=0;	% matrix to hold segments of raw 
										% data corresponding to each value
										% of alpha and each measurement

    beta(numOfTest)=0;					% vector to hold heating rates -- what for?

	collectedoutput = {};
	
	for ii=1:numOfTest
		
		% data now contains alpha between the temperature limits given in
		% the configuration file
        data = preProcessing(rawData{ii}, params.preProcessingParameters, mass(ii));

% 		makefig1([data(:,1) data(:,3)] );	% alpha vs. Temp

        b = polyfit(data(:,2),data(:,1),1);	% calculate approx. heating rate -- what for?
        beta(ii) = b(1);
		
		% Get the raw data segmented between the wanted values of alpha.
		% The first column of "res" contains cells that contain the raw
		% data between alpha(n-1) and alpha(n), including the interpolated
		% limits.
		% The second column of "res" contains the full measurement,
		% transposed from equally spaced time to equally spaced alpha.
		
% 		figure; hold on
% 		plot(data(:,1),data(:,3));
		
		res = getInterpedData(data, expectedAlpha);
        r=res{1};
        for jj=1:numOfRes
            rangeData{jj,ii}= r{jj};
        end
        interpData= res{2};
		
		interpData(:,4) = nan(size(interpData(:,1)));
		validRange = ~isnan(interpData(:,1));

% 		figure
% 		plot(interpData(validRange,1),gradient(interpData(validRange,3))./gradient(interpData(validRange,2)));
		
		firstP = find(validRange,1);
		lastP = find(validRange,1,'last');
		
		interpData(firstP:lastP,4) = gradient(interpData(validRange,3))./gradient(interpData(validRange,2));
				
        output(interpData, params.outputExlParams, ii);
		collectedoutput{end+1} = interpData; %#ok<AGROW>
		disp(['done interpolating ' char(params.inputDataParameters.files(ii))]);
% 		plot(data(:,1), data(:,3), 'b-', interpData(:,1), interpData(:,3), 'ro');
	end

    Eav = calculateActiveEnergies(rangeData, beta, params.boundOfEa, params.method);
    %xlswrite (params.outputfile, [expectedAlpha' Eav], params.sheetname, params.topleftCellName)
    Eava = cleanRes(Eav,expectedAlpha');    
    output(Eava, params.outputExlParams, numOfTest+1) % the last one is for Ea vs. alpha
% 	makefig1(collectedoutput);
%     makefig(Eava);
    
	Ea = Eava(:,2);
	alpha = Eava(:,1);
	J = Eava(:,4:3+numOfTest);
	dlogJdE = Eava(:,4+numOfTest:3+2*numOfTest);
	
	save('isoconv_results','Ea','alpha','J','dlogJdE');	
	
%	Jout = 
	1;
	
    
    function expectedAlpha= createExpectedProgressVector(params)
        expectedAlpha = params.startAlpha:params.delta:params.endAlpha;
    end

    function [data, mass]=dataLoading(inputDataParameters) 
        path= inputDataParameters.dir;
        test_number = inputDataParameters.num;
        files = inputDataParameters.files;
        data= cell(test_number,1);
        mass(test_number) = 0;
        for f=1:test_number
            filename=files{f};
            s = size(filename,2);
            ext=filename(s-2:s);
            if ( strcmpi( ext, 'xls') )
                sheetname=filename(1:s-4);
            data{f} = ExcelReader(path, filename, sheetname); % Acquiring data from excel files
            mass(f) =1;
            end % endif
        end  
        function [exlData, num] = ExcelReader(path,txt1,txt2) % Define function
            exl = actxserver('excel.application'); % Create a COM (Excel Component Object Model) server
            % After creating this server, it will be able to manipulate Excel objects as in the Excel VB scripts
            exlWkbk = exl.Workbooks; % Create a workbooks object using excel COM (Similar to )
            % exlFile = exlWkbk.Open('C:\Documents and Settings\KC\Desktop\Study Aboard\NJIT works\Master Project\Isoconversion Coding\Simulated-dsc-data\simulated-01Kmin.csv');
            exlFile = exlWkbk.Open([path txt1]); % Manipulate workbooks object, using "open" method to open an existing workbook
            exlSheet1 = exlFile.Sheets.Item(txt2); % Create Excel Sheet object to read specific sheets.
            eobj = exlSheet1.Columns.End(4); % Grab sheet information
            numrows = eobj.row; % Determine how much rows are there in the sheet
            dat_range = ['A2:C' num2str(numrows)]; % Define data range to read
            rng = exlSheet1.Range(dat_range); % Copy the sheet data
            num = numrows-1;
            exlData = rng.Value; % Read data value
            exlData = cell2mat(exlData); % Transfer value format
            exlWkbk.Close % Close workbooks object, descending objects will be terminate automatically
            exl.Quit % Close COM server
        end
    end

    function res = getInterpedData(data, interpedPoints)

		s1 = size(interpedPoints,2);	% number of values of alpha; 
										% same as numOfRes in main
        interpedData(s1,3)=0;
        ind = 1;
        cur = data(ind,:);				% first row of data
        L_rangeData{s1,1}=0;

		for i=1:s1
			
%			interpedData(i,:);			% output for debugging?
			
			% Temp: temperature from data file
			% time: time from data file
			% value: given alpha for which T and t are sought
			% ind1: row index where alpha(ind1) < value
			% ind2: row index where alpha(ind2) > value
            [Temp, time, ind1, ind2] = getInterped(data, interpedPoints(i));
			value = interpedPoints(i);
			
			if (cur(3) > value) || isnan(ind1)
								% This means there is no data between 
								% alpha(i) and alpha(i-1); therefore the
								% corresponding data segment will be empty,
								% and the transposed/interpolated data
								% array will be set to unphysical values.
				
                L_rangeData{i,1}= [];
%                 interpedData(i,:) = [-1, -1, -1];
                interpedData(i,:) = [NaN, NaN, NaN];
            else
                if (ind == 1)
                    L_rangeData{i,1}= [data(ind:ind1,:); Temp time value];
                else
                    L_rangeData{i,1}= [cur; data(ind:ind1,:); Temp time value];
                end
                cur = [Temp time value];
                interpedData(i,:) = cur;  
                ind= ind2;
			end
			res{1} = L_rangeData; %#ok<AGROW>
			res{2} = interpedData; %#ok<AGROW>

		end

        function [Temp, time, ind1, ind2] = getInterped(data, interpedPoint)
            [ind1, ind2, ind_alpha] = findIndices(data(:,3), interpedPoint);
%			Temp = interp1([data(ind1,3) data(ind2,3)],[data(ind1,1) data(ind2,1)],interpedPoint,'linear');
%			time = interp1([data(ind1,3) data(ind2,3)],[data(ind1,2) data(ind2,2)],interpedPoint,'linear');

			if isnan(ind1)
				Temp = NaN;
				time = NaN;
			else
				Temp = interp1(data(ind1:ind2,1),1 + ind_alpha - ind1);
				time = interp1(data(ind1:ind2,2),1 + ind_alpha - ind1);
			end
			
% 			plot(Temp,interpedPoint,'o');
% 			drawnow;

			ind1 = floor(ind_alpha);
			ind2 = ceil(ind_alpha);
			
%			hold on;
%			plot(data(ind1:ind2,1), data(ind1:ind2,3), 'b-', Temp, interpedPoint, 'ro');

            function [ind1, ind2, ind_alpha]= findIndices(alpha, interpedPoint)
				
				% We need to find the limits between which the value of
				% alpha is located.  The range we need to consider falls
				% between the first and last zero crossing of 
				% "alphaD" = (exp. alpha - given alpha)
				% In an ideal measurement there would be only one zero
				% crossing, but since the data is noisy, there might me
				% many.			
				% We do this by multiplying alphaD with alphaD shifted by
				% one element. If neighboring elements have different
				% signs, the product will be negative (same signs ->
				% positive).  Then we find the indices of the negative (or
				% zero) elements, and remember (one before) the first and
				% the last as the limits between which the actual zero
				% crossing is located.
				
                alphaD = alpha - interpedPoint;      %				
				s = length(alpha);
				
				alphaD_zero = find([alphaD(1); alphaD].*[alphaD; alphaD(s)]<=0);
				
				if isempty(alphaD_zero)
					ind1 = NaN;
					ind2 = NaN;
					ind_alpha = NaN;
 					disp(['   note: alpha=' num2str(interpedPoint) ' not found in file ' num2str(ii)]);
				else
					ind1 = alphaD_zero(1)-1;
					ind2 = alphaD_zero(length(alphaD_zero));
								
					% Now we fit a straight line to alphaD and find the index
					% where this line crosses zero.  This index will be used to
					% look up/interpolate time and temperature for
					% "InterpedPoint"

					ok = false;

					while not(ok)

						ind1 = ind1-1;
						ind2 = ind2+1;
						if ind1 < 1
							ind1 = 1;
						end
						if ind2 > s
							ind2 = s;
						end

						alphaD_noisy = alphaD(ind1:ind2);
						[p, ~, mu] = polyfit((ind1:ind2)', alphaD_noisy, 1);
						ind_alpha = roots(p)*mu(2)+mu(1);

						ok = p(1) > 0;		% must have positive slope of alpha vs. index
						ok = ok && (ind_alpha > ind1);	% root must be between
						ok = ok && (ind_alpha < ind2);	% ind1 and ind2

					end

% 					figure(12345)
% 					[y,delta] = polyval(p,ind1:ind2,S,mu);
% 					
% 					hold on;
% 					plot(ind1:ind2, alphaD_noisy(:)+interpedPoint, 'b-', ind_alpha, interpedPoint, 'ro');
% 					ind1;
				end
				
				
            end % END of function findIndices
        end % END of function getInterped
    end

    function treatedData = preProcessing(data, preProcessingParameters, mass)
        n = size(preProcessingParameters, 2);
        temp=data;

		for iii =1:n
			L_params = preProcessingParameters{iii};
			switch L_params.name
				case 'tempRange'
                    temp2 = trim(temp, L_params);
                case 'dsc2progress'
                     % HAVE NOT BEEN IMPLEMENTED 
                case 'tg2progress'
                     temp2 = getAlphaFromTG(temp, mass, L_params.maxInc);
			end
			clear temp;
			temp= temp2; 
			clear temp2;
		end
        treatedData = temp;

		function newData = trim(data, params)
            T = data(:,1);
            logic = (T >= params.boundary(1)) & ( T <= params.boundary(2));
            newData = [data(logic,1)  data(logic,2) data(logic,3)];
		end

		function Progress = getAlphaFromTG(tgData, mass, maxInc)
%			[minV,ind] = min(tgData(:, 3));
%			[maxV,indM] = max(tgData(:, 3));
			minV = 0; maxV = 1;
            Progress = [tgData(:,1)  tgData(:,2)  (tgData(:, 3)-minV)/ mass / maxInc];  
        end
    end

    function res = calculateActiveEnergies(rangeData, beta, params, method)
        numOfRes = size(rangeData,1);
        numhr = size(rangeData,2);
        res(numOfRes,1)=0;
		res(numOfRes,2*numhr+2)=0;
        for i=1:numOfRes
            mins=0;
            rr{numhr}=[]; % 0;
            if numhr > 1
                rr{1}=rangeData{i,1};
                mins=size(rangeData{i,1},1);
            end
            for m = 2:numhr % following 
                rr{m}=rangeData{i,m};
                s = size(rangeData{i,m},1);
                if ( s < mins)
                    mins =s;
                end
            end
            %mins = checkMinNums(rangeData{i,:});
            if ( mins <2) % this is a unusual case. we cannot calculate Ea
                res(i,1) = -100000;
            else
                
                res(i,:) =calculateActiveEnergy(rr, beta, params, method);
            end
            clear rr;
        end
	end

    function minNumbers=checkMinNums(Ttas)
        numhr = size(Ttas,2);
        if numhr <1
            minNumbers=0;
        else
            minNumbers=size(Ttas{1},1);
        end
        for m = 2:numhr % following 
            s = size(Ttas{m},1);
            if ( s < minNumbers)
                minNumbers =s;
            end
        end
    end

    function [Ea] = calculateActiveEnergy(Tta, beta, params, method)
        % passed-in Tta is a slice of interpData(i,:,:)
        E1 = params(1);      % Initialize searching lower bound (for smaller phi)
        E2 = params(2);      % Initialize searching upper bound (for larger phi)
        precision = params(3);
        delta = 10000 ;  % Initialize variable searching interval
        Eg1 = E1;      % Initialize searching lower bound of range
        Eg2 = E1+delta; % Initialize searching upper bound of range
        sf = 1;        % Searching flag, determine the searching direction
%         while delta >= precision
%             if Eg1 >= E1 && Eg2 >= E1 && Eg1 <= E2 && Eg2 <= E2
% 				sJ1 = sumJ(method,Eg1,Tta,beta);
% 				sJ2 = sumJ(method,Eg2,Tta,beta);
%                 if sJ1(1)<sJ2(1)
%                 % This condition will be activated once the trend change
%                     sf = sf*(-1); % Reverse searching direction
%                     if delta >= precision % A condition to exit the loop without further modification on activation energy
%                         delta = delta/10; % Shrink the searching interval so that searching will be done within the region just found
%                         Eg1 = Eg2; % Redefine upper bound
%                         Eg2 = Eg1+sf*delta; % Redefine lower bound
%                     end
%                 else % Continue searching
%                     Eg1 = Eg1+sf*delta; % Define next searching step
%                     Eg2 = Eg2+sf*delta; % Define next searching step
%                 end
%             else % out of the range [E1 -- E2]
%                 Eg1 = -100000;
%                 break            
%             end
%         end

		sJ = @(E) sumJonly(method,E,Tta,beta);
		
		
		[Eg1, Eg1val, exitflag ] = fminbnd(sJ,E1,E2,optimset('TolX',1e-12));


		
		Jhigh = sumJ(method,Eg1+100,Tta,beta);
		logJhigh = log10(Jhigh(2:end));
		Jlow = sumJ(method,Eg1-100,Tta,beta);
		logJlow = log10(Jlow(2:end));
		dlogJdE = (logJhigh-logJlow)/200;
		
		Ea = [Eg1 sumJ(method,Eg1,Tta,beta) dlogJdE];

	end % END of function calculateActiveEnergy
 
    % Minimization function
    function [Jsum] = sumJ(method,Eg,Tta,beta)
        switch method
            case 'Vyazovkin1997'
                Jsum = sumJ1997(Eg,Tta,beta);
            case 'Vyazovkin2000'
                Jsum= sumJ2000(Eg,Tta);
        end
	end

	function [Jsum] = sumJonly(method,Eg,Tta,beta)
		tmp = sumJ(method,Eg,Tta,beta);
		Jsum = tmp(1);
	end

    function [Jsum] = sumJ1997(Eg,Tts,beta)
%         R = 8.31451; % Universal gas constant, J/mol-K
%         numhr = size(Tta,2);
%         for m = 1:numhr %Construct T0 matrix
%             Tt = Tts{m};
%             s = size(Tt,1);
%             T0(m) = Tt(1, 1);
%             J(m) = sum((exp(-Eg/R./Ta{m}(2:size(Ta{m},2)))+exp(-Eg/R./Ta{m}(1:size(Ta{m},2)-1)))/2.*(ta{m}(2:size(ta{m},2))-ta{m}(1:size(ta{m},2)-1)));
%         end
%         x = Eg/R./T0;
%         j0 = beta.^-1*(Eg/R).*exp(-x)./x.*(x.^2+10*x+18)./(x.^3+12*x.^2+36*x+24);
%         for i =1:numhr
%         end
%         Jsum = 0;
%         for i = 1:numhr
%             j1 = j0(i)+J(i);
%             for j = 1:numhr
%                 if i ~= j
%                     j2 = j0(j)+J(j);
%                     if j2 == 0
%                          i
%                          j
%                          Eg
%                     end
%                     Jsum = Jsum + j1/j2;
%                 end
%             end
%         end
    end
    function [Jsum] = sumJ2000(Eg,Tts)
        R = 8.31451; % Universal gas constant, J/mol-K
        numhr = size(Tts,2);
%        J(numhr)=0;
		J = zeros(size(Tts));
        for m = 1:numhr %Construct T0 matrix
            Tt = Tts{m};
            s = size(Tt,1);
            J(m) = sum(		...   
		(exp(-Eg/R./Tt(2:s,1)) + exp(-Eg/R./Tt(1:s-1,1)) ) / 2.*(Tt(2:s,2) - Tt(1:s-1,2)));
        end
        Js = 0;
		for i = 1:numhr
            j1 = J(i);
            for j = 1:numhr
                if i ~= j
                    j2 = J(j);
                    Js = Js + j1/j2;
                end
            end
		end
		Js = Js - (numhr^2-numhr);
		Jsum = [Js J];
    end

    function cleanedEav =cleanRes(Eav,aa)
        %Eav and aa should be same size vertical vector
		s = size(Eav);
        b = (Eav(:,1) > 0); % bad data will be -100000 so positive values are valid
        cleanedEav(:,1)=aa(b);
        cleanedEav(:,2:s(2)+1)=Eav(b,:);
    end

    function makefig(Eav)
        figure;
        plot(Eav(:,1),Eav(:,2)/1000,'ro',Eav(:,1),Eav(:,2)/1000,'b-');
		xlabel('Reaction progress');
		ylabel('Activation Energy (kJ/mol)');
		grid on;
		hold on;
%         title('Ea as function of degree of conversion')
        %set(gca,'XTick',0:0.05:1,'YTick',-100:50:500);
    end

    function makefig1(data)
        figure;
		for i = 1:length(data)
			x = data{i}(:,1);
			y = data{i}(:,3);
			x(x<0) = [];
			y(y<0) = [];
			plot(x,y,'ro',x,y,'b-'); hold on;
		end
		xlabel('Temperature (K)');
		ylabel('Degree of conversion');
		grid on;
		hold on;
%         title('Progress as function of temperature')
%         set(gca,'XTick',200:100:1700,'YTick',-.1:.1:1);
    end

    function output(data, outputXlsParams, n)
        outputfile = outputXlsParams.file;
        sheetname = outputXlsParams.sheetLocation{n,1};
        topleftCellName = outputXlsParams.sheetLocation{n,2};
        xlswrite (outputfile, data, sheetname, topleftCellName)        
    end

end
