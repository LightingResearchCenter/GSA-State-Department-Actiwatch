function varargout = convertActiwatchData
%CONVERTACTIWATCHDATA Summary of this function goes here
%   Detailed explanation goes here

% Map file paths
timestamp = datestr(now,'yyyy-mm-dd_HHMM');
project = '\\ROOT\projects\GSA_Daysimeter\StateDepartment_2017\Actigraph_Data';
dbPath  = fullfile(project,[timestamp,'.mat']);
excelPath = fullfile(project,'GSA State Dept actigraph data.xlsx');


% Preallocate dataArray
dataArray = struct;

% Find the sheets we want
[~,sheets]	= xlsfinfo(excelPath);
sheets = sheets(contains(sheets,'raw'))';
subjects = regexprep(sheets,'(\d\d\d) raw','$1');

% Setup wait bar
nSheets = numel(sheets);
h = waitbar(0, ['Please wait processing sheet 0 of ',num2str(nSheets)]);

% Iterate through sheets
for iSheet = 1:numel(sheets)
    waitbar(iSheet/nSheets, h, ['Please wait processing sheet ',num2str(iSheet),' of ',num2str(nSheets)]);
    
    thisSheet   = sheets{iSheet};
    thisSubject = subjects{iSheet};
    calcSheet   = [thisSubject,' calc'];
    
    % Read data from file
    data             = importActiwatchExcel(excelPath,thisSheet);
    [sleepB, sleepI] = importActiwatchCalc( excelPath,calcSheet);
    
    
    % Determine dates of baseline and intervention
    weeks = dates2weeks(data.DateTime);
    idxB = cellfun(@(x)any(ismember(sleepB.startDates,x)), weeks);
    idxI = cellfun(@(x)any(ismember(sleepI.startDates,x)), weeks);
    baselineDates     = weeks{idxB};
    interventionDates = weeks{idxI};
    
    % Limit bounds to dates worn
    compliance = ~(isnan(data.Activity) | strcmp(data.IntervalStatus,'EXCLUDED'));
    % Find unique dates, remove time component
    unqDates = unique(dateshift(data.DateTime(compliance),'start','day'));
    % Baseline bounds
    baselineIntersection = unqDates(ismember(unqDates, baselineDates));
    startBaseline = min(baselineIntersection);
    endBaseline   = dateshift(max(baselineIntersection),'end','day');
    % Intervention bounds
    interventionIntersection = unqDates(ismember(unqDates, interventionDates));
    startIntervention = min(interventionIntersection);
    endIntervention   = dateshift(max(interventionIntersection),'end','day');
    
    
    % Save to data array
    % Baseline
    data.Observation = data.DateTime >= startBaseline & data.DateTime <= endBaseline;
    data.Compliance  = compliance & data.Observation;
    
    dataArray(iSheet*2 - 1).subject = thisSubject;
    dataArray(iSheet*2 - 1).session = 'Baseline';
    dataArray(iSheet*2 - 1).data = data;
    
    % Intervention
    data.Observation = data.DateTime >= startIntervention & data.DateTime <= endIntervention;
    data.Compliance  = compliance & data.Observation;
    
    dataArray(iSheet*2 - 1).subject = thisSubject;
    dataArray(iSheet*2 - 1).session = 'Intervention';
    dataArray(iSheet*2 - 1).data = data;
end

delete(h);

save(dbPath,'dataArray');

if nargout > 0
    varargout{1} = dataArray;
end

end



function weeks = dates2weeks(dates)
    
    % Find unique dates, remove time component
    unqDates = unique(dateshift(dates,'start','day'));
    % Remove weekends
    weekdays = unqDates(~isweekend(unqDates));
    % Find indices where weeks break
    idxBreak = find(diff(weekday(weekdays)) < 0);
    % Throw a warning and short circuit if there are not enough weeks
    if isempty(idxBreak)
        warning('less than two weeks')
        weeks = {weekdays};
    else
        nWeeks = numel(idxBreak) + 1;
        weeks = cell(nWeeks,1);
        for iWeek = 1:nWeeks
            if iWeek == 1
                weeks{iWeek} = weekdays(1:idxBreak(1));
            elseif iWeek == nWeeks
                weeks{iWeek} = weekdays(idxBreak(end)+1:end);
            else
                weeks{iWeek} = weekdays(idxBreak(iWeek-1)+1:idxBreak(iWeek));
            end
        end
    end
end
