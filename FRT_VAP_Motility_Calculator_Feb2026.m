% ----------------------------------------------------------
% CASA METRICS CALCULATION — BATCH + COMBINED EXPORT
% ----------------------------------------------------------

clear; clc;

% --- Select folder ---
inputDir = uigetdir(pwd, 'Select folder with sperm tracking files');
if inputDir == 0
    error('No folder selected.');
end

% --- Find Excel and CSV files ---
files = [ ...
    dir(fullfile(inputDir, '*.xlsx')); ...
    dir(fullfile(inputDir, '*.csv')) ...
];

if isempty(files)
    error('No Excel or CSV files found.');
end

% --- Parameters ---
windowSize = 5;
minTime = 0.3;

% --- Master tables ---
masterTracks  = table();
masterSummary = table();

% ==========================================================
% LOOP THROUGH FILES
% ==========================================================
for f = 1:length(files)

    filename = files(f).name;
    filepath = files(f).folder;
    [~, baseName, ~] = fileparts(filename);

    fprintf('\nProcessing: %s\n', filename);

    data = readtable(fullfile(filepath, filename));

    ids = data.TRACK_ID;
    x   = data.POSITION_X;
    y   = data.POSITION_Y;
    t   = data.POSITION_T;

    trackIDs = unique(ids);
    results = {};

    % ======================================================
    % TRACK CALCULATIONS
    % ======================================================
    for k = 1:length(trackIDs)

        trk = trackIDs(k);
        mask = ids == trk;

        X = x(mask);
        Y = y(mask);
        T = t(mask);

        if isempty(X) || (max(T) - min(T)) < minTime
            continue
        end

        [T, order] = sort(T);
        X = X(order);
        Y = Y(order);

        dX = diff(X);
        dY = diff(Y);
        segDist = sqrt(dX.^2 + dY.^2);

        totalDist = sum(segDist);
        displacement = sqrt((X(end)-X(1))^2 + (Y(end)-Y(1))^2);

        duration = T(end) - T(1);
        VCL = totalDist / duration;
        VSL = displacement / duration;
        LIN = VSL / VCL;

        smaX = movmean(X, windowSize);
        smaY = movmean(Y, windowSize);

        vapDX = diff(smaX);
        vapDY = diff(smaY);
        vapDist = sqrt(vapDX.^2 + vapDY.^2);

        VAP = sum(vapDist) / duration;

        STR = VSL / VAP;
        WOB = VAP / VCL;

        results(end+1,:) = {
            trk, VCL, VSL, VAP, totalDist, displacement, LIN, STR, WOB
        };
    end

    if isempty(results)
        continue
    end

    % ======================================================
    % RESULTS TABLE
    % ======================================================
    resultsTable = cell2table(results, ...
        'VariableNames', {'TrackID','VCL','VSL','VAP','TotalDist', ...
                          'Displacement','Linearity','Straightness','Wobble'});

    % 🔹 Add source file column label
    resultsTable.SourceFile = repmat(string(baseName), height(resultsTable), 1);
    resultsTable = movevars(resultsTable, 'SourceFile', 'Before', 1);

    % ======================================================
    % MOTILITY + HYPERACTIVATION
    % ======================================================
    n = height(resultsTable);
    motility = strings(n,1);
    hyperactive = zeros(n,1);

    for i = 1:n
        VAP_i = resultsTable.VAP(i);
        VCL_i = resultsTable.VCL(i);
        LIN_i = resultsTable.Linearity(i);

        if LIN_i < 0.5 && VCL_i > 100
            hyperactive(i) = 1;
        end

        if VAP_i < 1
            motility(i) = "Immotile";
        elseif VAP_i >= 1 && VAP_i < 5
            motility(i) = "Non-progressive";
        elseif VAP_i >= 5 && VAP_i < 25
            motility(i) = "Medium Progressive";
        else
            motility(i) = "Rapid Progressive";
        end
    end

    resultsTable.Motility = motility;
    resultsTable.Hyperactivated = hyperactive;

    % ======================================================
    % SUMMARY TABLE
    % ======================================================
    totalTracks = height(resultsTable);

    categories = ["Rapid Progressive","Medium Progressive","Non-progressive","Immotile"];
    counts = zeros(numel(categories),1);

    for i = 1:numel(categories)
        counts(i) = sum(resultsTable.Motility == categories(i));
    end

    percentages = (counts / totalTracks) * 100;

    motilitySummary = table(categories', counts, percentages, ...
        'VariableNames', {'MotilityCategory','Count','Percentage'});

    hyperCount = sum(resultsTable.Hyperactivated == 1);
    hyperPercent = (hyperCount / totalTracks) * 100;

    hyperSummary = table("Hyperactivated", hyperCount, hyperPercent, ...
        'VariableNames', {'MotilityCategory','Count','Percentage'});

    finalSummary = [motilitySummary; hyperSummary];

    % Add source file label to summary
    finalSummary.SourceFile = repmat(string(baseName), height(finalSummary), 1);
    finalSummary = movevars(finalSummary, 'SourceFile', 'Before', 1);

    % ======================================================
    % SAVE INDIVIDUAL FILE
    % ======================================================
    outputFile = fullfile(filepath, ['results_' baseName '.xlsx']);
    writetable(resultsTable, outputFile, 'Sheet', 'Tracks');
    writetable(finalSummary, outputFile, 'Sheet', 'Summary');

    % Append to master
    masterTracks  = [masterTracks; resultsTable];
    masterSummary = [masterSummary; finalSummary];

end

% ==========================================================
% SAVE MASTER FILE
% ==========================================================
combinedFile = fullfile(inputDir, 'ExperimentID_Motility_Results.xlsx');

writetable(masterTracks, combinedFile, 'Sheet', 'All_Tracks');
writetable(masterSummary, combinedFile, 'Sheet', 'All_Summaries');

fprintf('\nBatch CASA processing COMPLETE.\n');

% ==========================================================
% GROUP BY TIMEPOINT + CHIP
% ==========================================================

fprintf('\nCreating Timepoint + Chip summary...\n');

allTracks = masterTracks;

nRows = height(allTracks);
timepoints = strings(nRows,1);
chips = strings(nRows,1);

for i = 1:nRows

    name = allTracks.SourceFile(i);
% --- Extract Timepoint (xhr) ---
tp = regexp(name, 'TP\d+', 'match');
if ~isempty(tp)
    timepoints(i) = tp{end};
else
    timepoints(i) = "Unknown";
end


    % --- Extract Chip (C#) ---
    chipMatch = regexp(name, 'C\d+', 'match');
    if ~isempty(chipMatch)
        chips(i) = chipMatch{end};  % Uses last c
    else
        chips(i) = "Unknown";
    end
end

allTracks.Timepoint = timepoints;
allTracks.Chip = chips;

% ==========================================================
% MEAN CASA PARAMETERS (per Timepoint + Chip)
% ==========================================================

numericVars = {'VCL','VSL','VAP','TotalDist','Displacement', ...
               'Linearity','Straightness','Wobble'};

meanTable = groupsummary(allTracks, ...
                         {'Timepoint','Chip'}, ...
                         'mean', numericVars);

% ==========================================================
% MOTILITY COUNTS (per Timepoint + Chip)
% ==========================================================

categories = ["Rapid Progressive","Medium Progressive", ...
              "Non-progressive","Immotile"];

uniqueGroups = unique(allTracks(:,{'Timepoint','Chip'}));

motilityResults = [];

for i = 1:height(uniqueGroups)

    tp = uniqueGroups.Timepoint(i);
    chip = uniqueGroups.Chip(i);

    mask = allTracks.Timepoint == tp & ...
           allTracks.Chip == chip;

    subset = allTracks(mask,:);

    counts = zeros(1,length(categories));

    for c = 1:length(categories)
        counts(c) = sum(subset.Motility == categories(c));
    end

    hyperCount = sum(subset.Hyperactivated == 1);

    motilityResults = [motilityResults;
                       tp, chip, counts, hyperCount];
end

%Convert numeric part separately to avoid cell issues
numericPart= str2double(motilityResults(:,3:end));

motilityTable = array2table(numericPart);
motilityTable.Properties.VariableNames = ...
    ["Rapid","Medium","NonProg","Immotile","Hyperactivated"];

%Add grouping columns separately

motilityTable.Timepoint = string(motilityResults(:,1));
motilityTable.Chip = string(motilityResults(:,2));

motilityTable = movevars(motilityTable, ...
                         {'Timepoint','Chip'}, 'Before', 1);

% ==========================================================
% MERGE MEANS + COUNTS
% ==========================================================

finalGrouped = outerjoin(meanTable, motilityTable, ...
                         'Keys',{'Timepoint','Chip'}, ...
                         'MergeKeys',true);
% ==========================================================
% CALCULATE PERCENTAGES PER CHIP
% ==========================================================

% Total sperm per chip (sum of all categories)
totalSperm = finalGrouped.Rapid + ...
             finalGrouped.Medium + ...
             finalGrouped.NonProg + ...
             finalGrouped.Immotile;

% Progressive sperm (exclude immotile)
progressiveTotal = finalGrouped.Rapid + ...
                   finalGrouped.Medium + ...
                   finalGrouped.NonProg;

% Motility percentages
finalGrouped.Rapid_pct    = (finalGrouped.Rapid    ./ totalSperm) * 100;
finalGrouped.Medium_pct   = (finalGrouped.Medium   ./ totalSperm) * 100;
finalGrouped.NonProg_pct  = (finalGrouped.NonProg  ./ totalSperm) * 100;
finalGrouped.Immotile_pct = (finalGrouped.Immotile ./ totalSperm) * 100;
finalGrouped.Total_Motility_pct = (progressiveTotal ./ totalSperm) * 100;

% Hyperactivation percentage (relative to progressive sperm only)
finalGrouped.Hyper_pct = (finalGrouped.Hyperactivated ./ progressiveTotal) * 100;

% ==========================================================
% SAVE TO MASTER FILE
% ==========================================================

writetable(finalGrouped, combinedFile, ...
           'Sheet','Grouped_By_Timepoint_Chip');

fprintf('Timepoint + Chip summary saved.\n');
