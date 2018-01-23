function anylizeActiwatchData
%ANYLIZEACTIWATCHDATA Summary of this function goes here
%   Detailed explanation goes here

% Enable dependencies
[githubDir,~,~] = fileparts(pwd);
circadianDir = fullfile(githubDir,'circadian');
addpath(circadianDir);


% Map paths
timestamp = datestr(now,'yyyy-mm-dd_HHMM');

projectDir = '\\ROOT\projects\GSA_Daysimeter\StateDepartment_2017\Actigraph_Data';

ls = dir([projectDir,filesep,'*.mat']);
[~,idxMostRecent] = max(vertcat(ls.datenum));
dataName = ls(idxMostRecent).name;
dataPath = fullfile(projectDir,dataName);

xlsxPath = fullfile(projectDir,[timestamp,'_ActiwatchAnalyses.xlsx']);


% Import source data
load(dataPath);


% Initialize output
T = table;
T.subject = vertcat({dataArray.subject})';
T.session = vertcat({dataArray.session})';
% Perform analysis
for iD = 1:numel(dataArray)
    % Extract data
    data    = dataArray(iD).data;
    idxKeep = data.Observation & data.Compliance;
    data    = data(idxKeep,:);
    % Shortcircuit if no useable data
    if isempty(data)
        warning(['Subject ',T.subject{iT},' ',T.session{iT},' is empty.'])
        continue
    end
    
    % Find samples belonging to the last day
    lastTime = max(data.DateTime);
    idxLast  = data.DateTime >= dateshift(lastTime,'start','day');
    
    % Perform IS and IV analysis
    [IS_all, IV_all ] = isiv2(data.DateTime,          data.Activity         );
    [IS_last,IV_last] = isiv2(data.DateTime(idxLast), data.Activity(idxLast));
    
    % Perform cosinor analysis
    [~, ~, phi_all ] = phasor.cosinorfit(datenum(data.DateTime), data.Activity, 1, 1);
    [~, ~, phi_last] = phasor.cosinorfit(datenum(data.DateTime(idxLast)), data.Activity(idxLast), 1, 1);
    % Convert radians to time of day
    acrophase_all  = duration(mod( phi_all,2*pi)*12/pi, 0, 0);
    acrophase_last = duration(mod(phi_last,2*pi)*12/pi, 0, 0);
    
    % Find the number of days used
    nHours_all  = hours(numel(data.DateTime)*mode(diff(data.DateTime)));
    nHours_last = hours(numel(data.DateTime(idxLast))*mode(diff(data.DateTime)));
    
    % Assign results to table
    T.nHours_all(iD)     = nHours_all;
    T.nHours_last(iD)    = nHours_last;
    T.IS_all(iD)         = IS_all;
    T.IS_last(iD)        = IS_last;
    T.IV_all(iD)         = IV_all;
    T.IV_last(iD)        = IV_last;
    T.acrophase_all(iD)  = acrophase_all;
    T.acrophase_last(iD) = acrophase_last;
    
end % end of for

T = sortrows(T);

% Save results
writetable(T, xlsxPath);
winopen(xlsxPath);

end



