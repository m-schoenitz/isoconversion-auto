clc
listing = dir;
ignorethese= {'.','..','Baseline','Vyazovkin 2000 calculator-v2','data'};

%% User adjustable parameters

% To use a new set of measurement files, run `clear all'

% This is the temperature range where all measurements should
% match/coincide/be the same. Consider this the initial state.

fitlimits = [370, 380]; 

% This is the minimum mass you want to use for Ea

m_start = 0;

% maximum mass to be used for Ea 

m_final = 0.22;


% mass increment (should probably not be smaller than 0.01)

m_increment = 0.01;

% max temperature where to consider data valid
% Note that the minimum temperature is the same as fitlimits(1)

T_max = 1200; % in K

% for plotting only, you can ignore this
offset = 0.0; 
offmult = 0; 

%% Here we get a folder name (either we have to ask for a new one, 
% or we use the one that we used previously)

try
	
	if ~isa(fo,'char')
		error('fo');
	end
	disp('using set of files that is already loaded...')
	
catch
	disp('loading new set of files...')
	fo = uigetdir('./','DATA MUST BE IN A SUBFOLDER FROM HERE');
	
	cur = pwd;
	fo = fo(length(cur)+2:end);
	
end

clear('alldata');

out_data=[];
Ealegend = {};
% Eatable = table;

warning('off', 'MATLAB:MKDIR:DirectoryExists');

figure(999); clf; figure(888); clf;


% for i = 1:length(listing)

% 	fo = listing(i).name;
sublisting = dir(fo);

common = [];
measurement_count = 0;

%% Here we read in all data files that start with "ExpDat_"

jdx = ~startsWith({sublisting.name},'ExpDat_');
sublisting(jdx) = [];

for j = 1:length(sublisting)
	
	fi = sublisting(j).name;
	tgdata = loadDSCTG([fo,'/',fi]);
	
	if ~exist('alldata','var')
		alldata = tgdata;
	else
		alldata(end+1) = tgdata; %#ok<SAGROW>
	end
	T = tgdata.T;
	m = tgdata.m;
	
	fitrange = ((T > fitlimits(1)) & (T < fitlimits(2)));
	common = [common; [T(fitrange),m(fitrange)]]; %#ok<AGROW>
	measurement_count = measurement_count + 1;
end

%% 

if measurement_count > 1
	
	status = mkdir([fo,'/data/']);
	
	p = polyfit(common(:,1),common(:,2),1);
	
%% pre-processing of each data file?	
	
	figure(999);

	for k = 0:measurement_count-1
		
		disp(alldata(end-k).ID);
		
		
		T = alldata(end-k).T;
		m = alldata(end-k).m;
		t = alldata(end-k).t;
		
% 		figure(12345)
% 		plot(t,m);
% 		keyboard
		
		interprange = T > 100;
		fitrange = ((T > fitlimits(1)) & (T < fitlimits(2)));
		
		% 						m200 = interp1(smooth(T(interprange)),m(interprange),200);
		% 						m200 = interp1(T(interprange),m(interprange),200);
		%
		% 						m2 = m/m200-1;

		m_min = min(m);
		m2 = m/m_min;
				
		p = polyfit(T(fitrange),m2(fitrange),1);
		
		% 			m2 = (m2-polyval(p,T));
		
		m2 = (m2-polyval(p,fitlimits(1)));
		y=m2+offset*offmult;
		
		figure(999);
		
		plot(T,y);
		hold on;
		outname = ['data',num2str(k+1)];
		writetable(table([T+273.15,t,m2]), ...
			[fo,'/data/',outname], ...
			'filetype','spreadsheet', ...
			'sheet',outname, ...
			'WriteVariableNames',false,...
			'useexcel',false);
		
		Ealegend{end+1} = sprintf('%3.1f:%s',alldata(end-k).HR,alldata(end-k).ID); %#ok<SAGROW>
		offmult = offmult+1;

	end
	
%% writing config_TGA
	
	configfile = write_config_TGA(fo, ...
		measurement_count, ...
		[m_start,m_final,m_increment], ...
		[fitlimits(1)+273.15,T_max]);

%% running the actual isoconversion
	
	Ea = isoconv(configfile, 'sheet1', 'A2:A100', 'B2');


%% plot the results
	
	figure(888);
	plot(Ea(:,1),Ea(:,2)/1000,'.-');
	hold on;
	
	% 		massrange = [false; Ea(2:end,1)<0.20];
	% 		[p,S,mu] = polyfit(Ea(massrange,1),Ea(massrange,2)/1000,1);
	% 		[Ea0,deltaEa0] = polyval(p,0,S,mu);
	%
	% 		Eatable(end+1,:) = table(Ealegend(end), Ea0, deltaEa0, 'variablenames',{'ID','Ea','dEa'}); %#ok<SAGROW>
	
end

% end

%% add axis labels and legends, etc.

figure(888);
% legend(Ealegend,'interpreter','none');
xlabel('(m-m_0)/m_0')
ylabel('Activation Energy, kJ/mol')
% ylim([50,350])

figure(999);
% legend(Ealegend(end:-1:1),'interpreter','none','location','northwest');
legend(Ealegend,'interpreter','none','location','northwest');
% ylim([-0.01,2]);
xlabel('T, °C');
ylabel('(m-m_0)/m_0, offset and grouped by material');

% Eatable %#ok<NOPTS>

