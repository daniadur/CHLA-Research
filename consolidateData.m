clc
clear
warning('off', 'all')
%% Import Ultrasound Data

% Prompt User to Select Excel File
fprintf('Please select an Excel file to be read:\n');
[ult_file, ult_path] = uigetfile('*.xlsx');
fprintf('Reading %s...\n \n', ult_file);

% Identify patient acrostic
if contains(ult_file, '-')
    patient_id = extractBefore(ult_file, '-');
elseif contains(ult_file, '_')
    patient_id = extractBefore(ult_file, '_');
end

% Identify sheet names
sheet1 = ['Data', ' - ', patient_id, '_Baseline']; 
sheet2 = ['Data', ' - ', patient_id, '_Deflation'];

% Create Spreadsheet Datastore
data_sheet = spreadsheetDatastore(fullfile(ult_path, ult_file)); % access spreadsheet data
sheet_names = sheetnames(data_sheet, data_sheet.Files{1});

% Find correct datasheet
for i=1:1:sheet_names.length 
    if contains(sheet_names(i), sheet1)
        base_sheet = sheet_names(i);
    end
    if contains(sheet_names(i), sheet2)
        defl_sheet = sheet_names(i);
    end
end

data_sheet.Sheets = base_sheet; % read in baseline data only
data_sheet.SelectedVariableNames = {'MSEC', 'BDIAMM'}; % reads these two columns
ult_baseline = read(data_sheet);
ult_baseline{1,1} = 0; % time info is missing, set to zero
ult_baseline{:,1} = ult_baseline{:,1}./1000; % converting msec to sec

data_sheet.Sheets = defl_sheet; % read in deflation data only
data_sheet.SelectedVariableNames = {'MSEC', 'BDIAMM'};
ult_deflation = read(data_sheet);
ult_deflation{1,1} = 0; % time info is missing use zero
ult_deflation{:,1} = ult_deflation{:,1}./ 1000; % converting msec to sec

%% Import PU Data
fprintf('Please select a MAT file to be read:\n');
[pu_file, pu_path] = uigetfile('*.mat', ult_path);
fprintf('Loading %s...\n \n', pu_file);

% Ultrasound Generated Time 
genBaseTime = '2018/05/05 11:02:16 AM';
genDeflTime = '2018/05/05 11:12:27 AM';

% Loading PU file
load(fullfile(pu_path,pu_file));
puTime = ((0:1:size(pu2)-1)/fs_high)'; % generating col time vector based on fs
puData = table(puTime,pu2, 'VariableNames', {'SEC', 'DATA'}); % create table
seg_time = cell2mat(event.sec(2)); % time difference
ind = find(contains(event.freeTxt_raw, "Segment 1")); % find time stamp
noted_time = extractAfter(event.freeTxt_raw(ind+1), ' '); % gets rid of day
index = cell2mat(strfind(noted_time,':')); % only keep time stamp
noted_time = extractBetween(noted_time, index(1)-2, index(2)+2); 

time = datevec(noted_time); % create time vectors from string
ult_tBase = datevec(extractBetween(genBaseTime, ' ', ' '));
ult_tDefl = datevec(extractBetween(genDeflTime, ' ', ' '));

%% Adjust Time
T = sum([0,0,0,3600, 60, 1].*time);
baseAdj = sum([0,0,0,3600, 60, 1].*ult_tBase) - T + seg_time;
deflAdj = sum([0,0,0,3600, 60, 1].*ult_tDefl) - T + seg_time;
%{
deltaBase = ult_tBase - time; % find time difference
baseAdj = sum([0,0,0,3600, 60, 1].*deltaBase) + seg_time; % convert dif to sec
deltaDefl = ult_tDefl - time;
deflAdj = sum([0,0,0,3600, 60, 1].*deltaDefl) + seg_time;
%}
%% Consolidate Data
fprintf('Consolidating data...\n');
ult_deflation{:,1} = ult_deflation{:,1} + deflAdj; % adjust times
ult_baseline{:,1} = ult_baseline{:,1} + baseAdj;

% Create time vectors based on fs low
baseMSEC = ult_baseline.MSEC(1):1/fs_low:ult_baseline.MSEC(length(ult_baseline.MSEC));
baseBDIAMM = (interp1(ult_baseline.MSEC, ult_baseline.BDIAMM, baseMSEC))'; % interpolate data
baseline = table(baseMSEC',baseBDIAMM, 'VariableNames', {'SEC', 'BDIAMM'}); % generate tables
defMSEC = ult_deflation.MSEC(1):1/fs_low:ult_deflation.MSEC(length(ult_deflation.MSEC));
defBDIAMM = (interp1(ult_deflation.MSEC, ult_deflation.BDIAMM, defMSEC))';
deflation = table(defMSEC',defBDIAMM, 'VariableNames', {'SEC', 'BDIAMM'});

%% Generate Combined Signal
timeVec = (0:1/fs_low:puData.SEC(length(puData.SEC)))'; % create new time vector
signalData = (linspace(NaN, NaN, length(timeVec)))';  % create empty signal vector
baseBegin = find(timeVec<=baseline.SEC(1), 1, 'last'); % find where baseline begins
baseEnd = find (timeVec<=baseline.SEC(length(baseline.SEC)), 1, 'last'); % baseline end time
deflBegin = find(timeVec <= deflation.SEC(1), 1, 'last'); % find closest defl beginning time
deflEnd = find(timeVec <= deflation.SEC(length(deflation.SEC)), 1, 'last'); % closest end time
signalData(baseBegin:baseEnd, :) = baseline.BDIAMM; % fill with bdiamm data
signalData(deflBegin:deflEnd, :) = deflation.BDIAMM; 

combinedUltra = table(timeVec, signalData, 'VariableNames', {'SEC', 'BDIAMM'});

%% Plot Diameter vs Time (sec)
fprintf('Generating graph...\n');
figure
hold on
xlabel('Time (sec)')
ylabel('Diameter')
title('Diameter vs. Time', 'Fontsize', 20)
plot(combinedUltra.SEC, combinedUltra.BDIAMM, '-b')
plot(puData.SEC, puData.DATA, '-r')
legend('Ultrasound', 'PU Signal')