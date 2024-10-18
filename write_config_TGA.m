function name = write_config_TGA(fo, nfiles, range, T)

% range argument: [lower limit, upper limit, stepsize] -- all values 0..1
% Temperature argument: array of [min. T, max. T], in K

% if nfiles > 7
% 	warning('More than 7 input files in auto-generated config_TGA. There will be unpredictable behavior.')
% end

fo = strrep(fo,'\','/');
% fp = [fo,'/'];
fp = ['./',fo,'/'];
% fn = [fo,'/config_TGA'];
fn = ['./',fo,'/config_TGA'];

row = 1;

out = cell(1);

row = row+1; out{row,1} = 'inputDataParameter';
row = row+1; out{row,1} = 'inputDataNum';	out{row,2} = nfiles;
row = row+1; out{row,1} = 'inputFileDir';	out{row,2} = [strrep(pwd(),'\','/'),'/',fo,'/data/'];

for i = 1:nfiles
	row = row+1; 
	out{row,1} = ['inputFile',num2str(i)];
	out{row,2} = ['data',num2str(i),'.xls'];
end

row = row+2; out{row,1} = 'outputParameter';
row = row+1; out{row,1} = 'outputFile';  out{row,2} = [fp,'output'];
row = row+1; out{row,1} = 'outputDataNum';	out{row,2} = nfiles+1;

for j = 1:nfiles
	row = row+1; 
	out{row,1} = ['data',num2str(j)];
	out{row,2} = 'sheet1';
	
	pad = '';
	if j > 7.5 % if 4*(j-1) > 26
		pad = char(65*(idivide(int8(4*(j-1)),26)));
	end
% 	out{row,3} = [pad,char(65+mod(4*(j-1),26)),'2'];
	out{row,3} = [pad,char(65+mod(5*(j-1),26)),'2'];
end

row = row+1; out{row,1} = 'sum';	out{row,2} = 'output';	out{row,3} = 'A2';

row = row+2; out{row,1} = 'progressParameters';
row = row+1; out{row,1} = 'startProg'; out{row,2} = range(1);
row = row+1; out{row,1} = 'endPro'; out{row,2} = range(2);
row = row+1; out{row,1} = 'stepSize'; out{row,2} = range(3);

row = row+2; out{row,1} = 'preProcessingParameters';
row = row+1; out{row,1} = 'steps';	out{row,2} = 2;
row = row+1; out{row,1} = 'tempRange';	out{row,2} = T(1); out{row,3} = T(2);
row = row+1; out{row,1} = 'tg2progress';	out{row,2} = 1;

disp(pwd)

writetable(table(out),fn,'filetype','spreadsheet','WriteVariableNames',false,'useexcel',false);
% writetable(table(out),fn);

name = [fn,'.xls'];

end
