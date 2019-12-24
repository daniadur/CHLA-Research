clc
clear

%% Prompt User to Select Excel File
fprintf('Please select an Excel file to be read:\n');
file = uigetfile('*.xlsx');

% identify patient acrostic
patient_id = extractBefore(file, '-');
%% Read in Data
% identify sheet names
sheet1 = ['Data', ' - ', patient_id, '_Baseline'];
sheet2 = ['1 Data', ' - ', patient_id, '_Deflation'];

data_sheet = spreadsheetDatastore(file);
data_sheet.Sheets = sheet1;
data_sheet.SelectedVariableNames = {'BDIAMM', 'MSEC'}; % reads these two columns
baseline = read(data_sheet);
baseline{1,2} = 0; % time info is missing set to zero
baseline{:,2} = baseline{:,2}./1000; % converting msec to sec

data_sheet.Sheets = sheet2;
data_sheet.SelectedVariableNames = {'BDIAMM', 'MSEC'};
deflation = read(data_sheet);
deflation{1,2} = 0; % time info is missing use zero
deflation{:,2} = deflation{:,2}./ 1000; % converting msec to sec

%% Plot Diameter vs Time (sec)
figure
hold on
xlabel('Time (sec)')
ylabel('Diameter')
title('Diameter vs. Time', 'Fontsize', 20)
plot(baseline.MSEC, baseline.BDIAMM) % plots table by columns
plot(deflation.MSEC, deflation.BDIAMM)
legend('Baseline', 'Dilation')