% ----------------------------------------------------------
% CASA METRICS CALCULATION — BATCH EXCEL PROCESSING
% Handles a folder of Excel/CSV files
% Saves results_<filename>.xlsx in same folder
% ----------------------------------------------------------

clear; clc;

% --- Select folder containing tracking files ---
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
    error('No Excel or CSV files found in selected folder.');
end

% --- Parameters ---
windowSize = 5; % SMA window for VAP path
minTime = 0.3;  % minimum track duration (sec)

% ==========================================================
% LOOP THROUGH FILES
% ==========================================================
for f = 1:length(files)

    filename = files(f).name;
    filepath = files(f).folder;

    fprintf('\nProcessing: %s\n', filename);

    % --- Load data ---
    data = readtable(fullfile(filepath, filename));

    % --- Rename for convenience ---
    ids = data.TRACK_ID;
    x   = data.POSITION_X;
    y   = data.POSITION_Y;
    t   = data.POSITION_T;

    trackIDs = unique(ids);
    results = {};

    % ======================================================
    % TRACK-LEVEL CALCULATIONS
    % ======================================================
    for k = 1:length(trackIDs)

        trk = trackIDs(k);
        mask = ids == trk;

        X = x(mask);
        Y = y(mask);
        T = t(mask);

        if isempty(X) || isempty(Y)
            continue
        end

        if (max(T) - min(T)) < minTime
            continue
        end

        % Sort by time
        [T, order] = sort(T);
        X = X(order);
        Y = Y(order);

        % --- Raw metrics ---
        dX = diff(X);
        dY = diff(Y);
        segDist = sqrt(dX.^2 + dY.^2);

        totalDist = sum(segDist);
        displacement = sqrt((X(end)-X(1))^2 + (Y(end)-Y(1))^2);

        duration = T(end) - T(1);
        VCL = totalDist / duration;
        VSL = displacement / duration;
        LIN = VSL / VCL;

        % --- VAP ---
        smaX = movmean(X, windowSize);
        smaY = movmean(Y, windowSize);

        vapDX = diff(smaX);
        vapDY = diff(smaY);
        vapDist = sqrt(vapDX.^2 + vapDY.^2);

        VAP = sum(vapDist) / duration;

        STR = VSL / VAP;
        WOB = VAP / VCL;

        % --- Store ---
        results(end+1,:) = {
            trk, VCL, VSL, VAP, totalDist, displacement, LIN, STR, WOB
        };
    end

    % ======================================================
    % RESULTS TABLE
    % ======================================================
    resultsTable = cell2table(results, ...
        'VariableNames', {'TrackID','VCL','VSL','VAP','TotalDist', ...
                          'Displacement','Linearity','Straightness','Wobble'});

    % --- Motility + hyperactivation ---
    n = height(resultsTable);
    motility = strings(n,1);
    hyperactive = zeros(n,1);

    for i = 1:n
        VAP_i = resultsTable.VAP(i);
        VCL_i = resultsTable.VCL(i);
        VSL_i = resultsTable.VSL(i);
        LIN_i = resultsTable.Linearity(i);

        if LIN_i < 0.5 && VCL_i > 100
            hyperactive(i) = 1;
        end

        if VAP_i < 5
            motility(i) = "Immotile";
        elseif VAP_i >= 5 && (VSL_i < 25 || LIN_i < 0.5)
            motility(i) = "Non-progressive";
        elseif LIN_i >= 0.5 && VAP_i >= 50
            motility(i) = "Rapid Progressive";
        elseif LIN_i >= 0.5 && VAP_i >= 25
            motility(i) = "Medium Progressive";
        else
            motility(i) = "Non-classified";
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

    % ======================================================
    % SAVE OUTPUT (same folder)
    % ======================================================
    [~, baseName, ~] = fileparts(filename);
    outputFile = fullfile(filepath, ['results_' baseName '.xlsx']);

    writetable(resultsTable, outputFile, 'Sheet', 'Tracks');
    writetable(finalSummary, outputFile, 'Sheet', 'Summary');

    fprintf('Saved: %s\n', outputFile);

end

fprintf('\nBatch CASA processing COMPLETE.\n');
