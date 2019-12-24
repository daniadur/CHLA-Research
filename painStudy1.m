%% Pain App Study

clc
clear
warning('off','MATLAB:table:ModifiedVarnames')
%% Prompt User

fprintf('Select pain study (Excel) file: \n');
[pfile,ppath] = uigetfile({'*.xlsx';'*.xls'}, 'Select pain study (Excel) file:');
%% Load File

% Create spreadsheet object
dataSheet = spreadsheetDatastore(fullfile(ppath,pfile));
dataSheet.SelectedVariableNames = {'Acrostic', 'Observed', 'AuraNow', 'PainNow'};
dataSheet.SelectedVariableTypes = {'char', 'datetime', 'char', 'char'};
allData = read(dataSheet);
oldData = allData;

% Scrub the data of 'No's and empty data
emptyAura = ~(contains(allData.AuraNow, 'Y'));
emptyPain = ~(contains(allData.PainNow, 'Y'));
empty = find(emptyAura == emptyPain);
allData(empty, :) = []; % only keeps yes's

% Save Dates and Times to strings
allDataDate = string(datetime(allData.Observed(:), 'Format','dd-MMM-yyyy'));
allDataTime = string(datetime(allData.Observed(:), 'Format','HH:mm:ss'));

patientnum = 1;
begin = 1;

fprintf('Generating patient reports...\n');

% Iterate through entire data
for i=1:length(allDataTime)
    
    if i < length(allDataTime) 
        End = i-1;
    else
        End = length(allDataTime); 
    end % create start and end points for new patients
    
    % Compare names to determine end points
    if (~(strcmp(allData.Acrostic(begin),allData.Acrostic(i)))) ...
            || (i == length(allDataTime)) 
        
        % Generate separate reports for each patient
        Reports{patientnum} = table(allData.Acrostic(begin:End), ...
            allDataDate(begin:End), allDataTime(begin:End),...
            allData.AuraNow(begin:End), allData.PainNow(begin:End), ...
            'VariableNames', {'Name', 'Date', 'Time', 'Aura', 'Pain'});
        Analysis{patientnum} = struct('Name', allData.Acrostic(begin), ...
            'TotalDays',0, 'PainDays', 0,'NoPainDays', 0);
        Analysis{patientnum}.AuraButNoPain = datetime.empty;
        Analysis{patientnum}.PainButNoAura = datetime.empty;
        Analysis{patientnum}.AuraNowPainNow = datetime.empty;
        Analysis{patientnum}.PainAfterAura = table('Size',[0 3], ...
            'VariableTypes', {'datetime','datetime','duration'}, ...
            'VariableNames', {'Aura','Pain','Change'});
        patientnum = patientnum + 1;
        begin = i;
    end        
    clear name End;
end
patientnum = patientnum - 1;

%% Analyze Data

fprintf('Analyzing data...\n');
for i=1:patientnum % iterate through each patient
    endDate = 0;
    k = 1;
    for j=1:length(Reports{i}.Name) % iterate through complete report
        
        % iterate through complete report and stop when hits endDate
        if k == endDate+1 
            clear times;
            
            % find first date and last date for each patient
            times = find(contains(Reports{i}.Date, Reports{i}.Date(j)));
            beginDate = times(1);
            endDate = times(end);
            
            % if only one day
            if beginDate == endDate 
                if contains(Reports{i}.Aura(j), 'Y') % if there is aura
                    if contains(Reports{i}.Pain(j), 'Y') % Aura Now Pain now
                        dateString = strcat(Reports{i}.Date(j), " ", ...
                            Reports{i}.Time(j));
                        dateString = datetime(dateString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss');
                        % add date to end of AuraNowPainNow vector
                        End = length(Analysis{i}.AuraNowPainNow) + 1;
                        Analysis{i}.AuraNowPainNow(End) = dateString;
                        Analysis{i}.PainDays = Analysis{i}.PainDays + 1;
                        
                    else % aura but no pain
                        dateString = strcat(Reports{i}.Date(j), " ", ...
                            Reports{i}.Time(j));
                        dateString = datetime(dateString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss');
                        End = length(Analysis{i}.AuraButNoPain) + 1;
                        Analysis{i}.AuraButNoPain(End) = dateString;
                        Analysis{i}.NoPainDays = Analysis{i}.NoPainDays + 1;
                    end
                else % if no aura
                    if contains(Reports{i}.Pain(j), 'Y')
                        dateString = strcat(Reports{i}.Date(j), " ", ...
                            Reports{i}.Time(j));
                        dateString = datetime(dateString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss a');
                        End = length(Analysis{i}.PainButNoAura) + 1;
                        Analysis{i}.PainButNoAura(End) = dateString;
                    end
                end
                Analysis{i}.TotalDays = Analysis{i}.TotalDays + 1; % update total days
                k = endDate+1; % update k
            end
        else % for multiple dates
            auraFound = 0;
            painFound = 0;
            for n = beginDate:endDate
                
                if ~auraFound % create aura and pain flags
                    if contains(Reports{i}.Aura(n), 'Y')
                        auraFound = n;
                    end
                end
                
                if ~painFound             
                    if contains(Reports{i}.Pain(n), 'Y')
                        painFound = n;
                    end
                end
                
                if auraFound && painFound % if both pain and aura
                    if auraFound == painFound
                        dateString = strcat(Reports{i}.Date(n), " ", ...
                            Reports{i}.Time(n));
                        dateString = datetime(dateString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss');
                        End = length(Analysis{i}.AuraNowPainNow)+1;
                        Analysis{i}.AuraNowPainNow(End) = dateString;
                        Analysis{i}.PainDays = Analysis{i}.PainDays + 1;
                        painFound = 0; % change flag so it finds new pain
                        
                    else % if aura and pain doesnt match
                        auraString = strcat(Reports{i}.Date(auraFound), ...
                            " ",Reports{i}.Time(auraFound));
                        auraString = datetime(auraString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss');
                        painString = strcat(Reports{i}.Date(painFound), ...
                            " ",Reports{i}.Time(painFound));
                        painString = datetime(painString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss');
                        change = painString - auraString;
                        End = height(Analysis{i}.PainAfterAura)+1;
                        Analysis{i}.PainAfterAura.Aura(End) = auraString; 
                        Analysis{i}.PainAfterAura.Pain(End) = painString;
                        Analysis{i}.PainAfterAura.Change(End) = change;
                        painFound = 0; % change flag so it can find new pain time
                        
                        if n == endDate % update pain day at end of iteration
                            Analysis{i}.PainDays = Analysis{i}.PainDays + 1;
                        end
                    end
                    
                elseif auraFound && ~painFound % if aura but no pain
                    if n ~= endDate % if before end of day
                        if contains(Reports{i}.Aura(auraFound+1), 'Y') % if aura is next
                            dateString = strcat(Reports{i}.Date(n), ...
                                " ",Reports{i}.Time(n));
                            dateString = datetime(dateString, 'Format', ...
                                'dd-MMM-yyyy hh:mm:ss a');
                            End = length(Analysis{i}.AuraButNoPain)+1;
                            Analysis{i}.AuraButNoPain(End) = dateString;
                        end
                    else % if n does == endDate
                        dateString = strcat(Reports{i}.Date(n), ...
                            " ",Reports{i}.Time(n));
                        dateString = datetime(dateString, 'Format', ...
                            'dd-MMM-yyyy hh:mm:ss a');
                        End = length(Analysis{i}.AuraButNoPain)+1;
                        Analysis{i}.AuraButNoPain(End) = dateString;
                        auraFound = 0; % update aura flag so it can find another one
                    end
                    
                    if n == endDate
                        Analysis{i}.NoPainDays = Analysis{i}.NoPainDays + 1;
                    end
                    
                elseif ~auraFound && painFound % pain but no aura
                    dateString = strcat(Reports{i}.Date(n), " ",Reports{i}.Time(n));
                    dateString = datetime(dateString, 'Format', 'dd-MMM-yyyy hh:mm:ss a');
                    End = length(Analysis{i}.PainButNoAura)+1;
                    Analysis{i}.PainButNoAura(End) = dateString;
                    painFound = 0; % update pain flag so it find next one
                end  
            end
            
            k = endDate + 1; % update k
            Analysis{i}.TotalDays = Analysis{i}.TotalDays + 1; % update total days
        end
    end
    
    % change horizontal vectors to vertical vectors
    Analysis{i}.AuraButNoPain = Analysis{i}.AuraButNoPain';
    Analysis{i}.PainButNoAura = Analysis{i}.PainButNoAura';
    Analysis{i}.AuraNowPainNow = Analysis{i}.AuraNowPainNow';  
    
    clear beginDate endDate days;
    
    %% Generate Results
    
    fprintf('Saving reports...\n');
    patientLogistics = table(string(Analysis{i}.Name), Analysis{i}.TotalDays, ...
        Analysis{i}.PainDays, Analysis{i}.NoPainDays, 'VariableNames', {'Name', 'TotalDays',...
        'PainDays', 'NoPainDays'});
    
    filename = strcat(string(Analysis{i}.Name), "DataAnalysis.xlsx");
    writetable(patientLogistics, filename, 'Sheet', 'Patient-Logistics')
    writematrix(Analysis{i}.AuraButNoPain, filename, 'Sheet', 'Aura-But-No-Pain')
    writematrix(Analysis{i}.PainButNoAura, filename, 'Sheet', 'Pain-But-No-Aura')
    writematrix(Analysis{i}.AuraNowPainNow, filename, 'Sheet', 'Aura-Now-Pain-Now')
    writetable(Analysis{i}.PainAfterAura, filename, 'Sheet', 'Pain-After-Aura')
end
