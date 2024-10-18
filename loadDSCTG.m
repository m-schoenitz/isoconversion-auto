function o = loadDSCTG(file)
% [filename,path] = uigetfile('*.xls?','Select a File');
% ffn = [path,filename];

[~,~,raw] = xlsread(file);

firstrow = 40; % this may be wrong.

[~,fname,ext] = fileparts(file);

ext = lower(ext);

switch ext
	case '.xls'
		
		o.m0 = raw{19,2};
		o.ID = raw{5,2};
		o.file = raw{2,2};
		
		o.T = cell2mat(raw(firstrow:end, 1));
		o.t = cell2mat(raw(firstrow:end, 2));
		o.m = cell2mat(raw(firstrow:end, 4));

	case '.csv'
		
		disp(fname)
% 		disp(size(raw))
% 		disp(raw(1,:))

		try
			temp = str2double(split(raw{19},','));
			o.m0 = temp(2);
			temp = split(raw{5},',');
			o.ID = strtrim(temp{2});
			temp = split(raw{2},',');
			o.file = strtrim(temp{2});
		catch
			o.m0 = cell2mat(raw(19,2));
			o.ID = strtrim(cell2mat(raw(5,2)));
			o.file = strtrim(cell2mat(raw(2,2)));
		end

		try
			raw = str2double(split(raw(firstrow:end,:),','));
			o.T = raw(:, 1);
			o.t = raw(:, 2);
			o.m = raw(:, 4);
		catch
			o.T = cell2mat(raw(firstrow:end,1));
			o.t = cell2mat(raw(firstrow:end,2));
			o.m = cell2mat(raw(firstrow:end,4));
		end

end

o.HR = mean(gradient(o.T)./gradient(o.t));

1;
end
