function data = importActiwatchExcel(filePath,sheet)
%UNTITLED Summary of this function goes here
%   Detailed explanation goes here

% Read data from file
[~,~,raw] = xlsread(filePath,sheet,'','basic');

% Find marker indicating start of raw data
idxLine     = strcmp(raw(:,1),'Line');
idxDate     = strcmp(raw(:,2),'Date');
idxTime     = strcmp(raw(:,3),'Time');
idxStatus	= strcmp(raw(:,4),'Off-Wrist Status');
idxActivity	= strcmp(raw(:,5),'Activity');

idxDataStart = find(idxLine & idxDate & idxTime & idxStatus & idxActivity);

if isempty(idxDataStart)
    data = [];
    return;
end
% Extract just the epoch by epoch data
rawData = raw(idxDataStart(1):end,:);
rawData(2,:) = [];

% replace string NaN with NaN
idxStrNaN = strcmp(rawData,'NaN');
rawData(idxStrNaN) = {NaN};

% Keep only columns with labels
idxColLabels = cellfun(@ischar,rawData(1,:));
rawData(:,~idxColLabels) = [];

% Construct valid variable names
variableNames = matlab.lang.makeValidName(rawData(1,:),'ReplacementStyle','delete');

% Convert data to table
data = cell2table(rawData(2:end,:),'VariableNames',variableNames);

% Combine date and time
data.DateTime = datetime((data.Date + data.Time),'ConvertFrom','excel','TimeZone','local');

% Move DateTime to front
data = [data(:,end),data(:,1:end-1)];

% Remove redundant variables
data.Line = [];
data.Date = [];
data.Time = [];

end

