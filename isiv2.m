function [interdailyStability,intradailyVariability] = isiv2(datetimeArray,dataArray)
% ISIV2 
%	Returns the interdaily stability (IS) and the
%	intradaily variability (IV) statistic for time series A.
%   datetimeArray is an ordered column vector of datetimes
%	dataArray is a column vector of equal size to the timeArray.
%	given by scalar epochSec in units of seconds.
%	Converts data to hourly increments before calculating metric.

interdailyStability = NaN;
intradailyVariability = NaN;

% Find the most frequent sampling epoch in seconds
epochSec = seconds(mode(diff(datetimeArray)));

% Check that input is useable
n = numel(dataArray);


if epochSec>3600
    warning('Cannot compute statistic becasue time increment is larger than one hour');
    return;
end

% Convert to hourly data increments
relDuration = datetimeArray - datetimeArray(1); % Convert to relative time
relHours    = floor(hours(relDuration) + 1); % Convert to hours
unqHours    = unique(relHours);
nHours      = numel(unqHours);

for iHour = nHours:-1:1
    idx = relHours == unqHours(iHour);
    hourlyDataArray(iHour,1) = mean(dataArray(idx));
end

if (n < 48 || n*epochSec < 48*3600)
    warning('Cannot compute interdail stability statistic because time series is less than 48 hours');
    interdailyStability = nan;
else
    A_p = modifiedenright(unqHours,hourlyDataArray,24);
    interdailyStability = A_p^2/var(hourlyDataArray,1);
end

deltaHourly = diff(hourlyDataArray)./diff(unqHours);
intradailyVariability = sum(deltaHourly.^2)/((nHours-1)*var(hourlyDataArray,1));

end


function amplitudeArray = modifiedenright(timeArrayHrs,dataArray,testPeriodArray)
% MODIFIEDENRIGHTPERGRAM Calculates the Enright periodogram on column vector
%   dataArray using range of periods given by testPeriodArray. Formulation 
%   taken from Philip G. Sokolove adn Wayne N. Bushell, The Chi Square 
%   Periodogram: Its Utility for Analysis of Circadian Rhythms, J. Theor. 
%   Biol. (1978) Vol 72, pp 131-160.

nPeriods = numel(testPeriodArray); % noted as P in source paper
amplitudeArray = zeros(nPeriods,1);
for i1 = 1:nPeriods
    p = testPeriodArray(i1); % period currently being tested
    P = floor(p); % number of columns
    K = ceil(max(timeArrayHrs)/p); % number of rows
    
    X = cell(K,P); % preallocate the data matrix
    cellXbar_hp = cell(P,1); % preallocate column mean array
    
    % reshape the dataArray into X
    P_i = mod(timeArrayHrs,p) + 1; % column position
    K_i = floor(timeArrayHrs/p) + 1; % row position
    for i2 = 1:P % move through columns
        for i3 = 1:K % move through rows
            idx = P_i==i2 & K_i==i3;
            if any(idx)
                X{i3,i2} = dataArray(idx);
            end
        end
        cellXbar_hp{i2} = mean(cell2mat(X(:,i2))); % column mean
        if isnan(cellXbar_hp{i2})
            cellXbar_hp{i2} = [];
        end
    end
    
    Xbar_hp = cell2mat(cellXbar_hp);
    Xbar_p = mean(Xbar_hp); % mean of column mean
    
    amplitudeArray(i1) = sqrt((1/P)*sum((Xbar_hp-Xbar_p).^2));
end

end

