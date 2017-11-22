function [sleepB, sleepI] = importActiwatchCalc(filePath,sheet)
%UNTITLED Summary of this function goes here
%   Detailed explanation goes here

% Read data from file
[~,~,raw] = xlsread(filePath,sheet,'','basic');

% Find markers indicating start week1 and week2
idxWeek1 = find(strcmp(raw(:,1),'week1'));
idxWeek2 = find(strcmp(raw(:,1),'week2'));

% Select start and end dates
sleepStartB = raw(idxWeek1+2:idxWeek2-1,3);
sleepEndB   = raw(idxWeek1+2:idxWeek2-1,5);

sleepStartI = raw(idxWeek2+2:end,3);
sleepEndI   = raw(idxWeek2+2:end,5);

% Filter out emty and nonnumeric cells
fValid = @(x) isnumeric(x) & ~isnan(x);

sleepStartB = cell2mat(sleepStartB(cellfun(fValid, sleepStartB)));
sleepEndB   = cell2mat(sleepEndB(  cellfun(fValid, sleepEndB  )));
sleepStartI = cell2mat(sleepStartI(cellfun(fValid, sleepStartI)));
sleepEndI   = cell2mat(sleepEndI(  cellfun(fValid, sleepEndI  )));

% Convert Excel dates to datetime and store in structs
sleepB = struct;
sleepI = struct;

sleepB.startDates = datetime(sleepStartB,'ConvertFrom','excel','TimeZone','local');
sleepB.endDates   = datetime(sleepEndB,  'ConvertFrom','excel','TimeZone','local');

sleepI.startDates = datetime(sleepStartI,'ConvertFrom','excel','TimeZone','local');
sleepI.endDates   = datetime(sleepEndI,  'ConvertFrom','excel','TimeZone','local');
end

