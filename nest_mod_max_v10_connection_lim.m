% Add the bimat package to the path
% for work laptop
%addpath(genpath('/Users/emilypayne/Documents/bimat_packadge'));

% for mac
addpath(genpath('/Users/emilypayne/Documents/MATLAB/github_repo'));

% for epg desktop
%addpath(genpath('C:\Users\emp.payne\Documents\bimat_github_repo'));

% Read the bipartite matrix from the Excel file
matrix = readmatrix('9bus bipartite matrices.xlsx', 'Sheet', 'WSCC9 Grid Original Matrix All', 'Range', 'B2:P22');

% Initialize variables
[m, n] = size(matrix);  
outputData = [];  % Initialize outputData to store results

% Define max connections for specific row and column ranges
maxConnectionsRowRange1 = 7;  % Max 7 connections for rows 1-21
maxConnectionsRowRange2 = 0;  % Max 0 connections for rows 28-42
maxConnectionsRow40to42 = 0;  % Max 0 connections for rows 40-42
maxConnectionsColumns1to9 = 7;  % Max 7 connections for columns 1-9
maxConnectionsColumns13to15 = 2;  % Max 2 connections for columns 13-15

% Loop to add all possible connections one by one
while true
    bestNODF = -inf;
    bestPosition = [];

    % Find the best position to add a new connection
    for i = 1:m
        for j = 1:n
            if matrix(i, j) == 0
                % Determine the max connections allowed for the current row
                if i >= 1 && i <= 21
                    maxConnectionsPerRow = maxConnectionsRowRange1;  % Max 4 for rows 1-21
                elseif i >= 28 && i <= 39
                    maxConnectionsPerRow = maxConnectionsRowRange2;  % Max 9 for rows 28-39
                elseif i >= 40 && i <= 42
                    maxConnectionsPerRow = maxConnectionsRow40to42;  % Max 6 for rows 40-42
                else
                    maxConnectionsPerRow = 0;  % No limit for other rows
                end

                % Determine the max connections allowed for the current column
                if j >= 1 && j <= 10
                    maxConnectionsPerColumn = maxConnectionsColumns1to9;  % Max 10 for columns 1-9
                elseif j >= 13 && j <= 16
                    maxConnectionsPerColumn = maxConnectionsColumns13to15;  % Max 8 for columns 13-15
                else
                    maxConnectionsPerColumn = 3;  % No limit for other columns
                end

                % Check row and column constraints before adding the connection
                if sum(matrix(i, :)) < maxConnectionsPerRow && sum(matrix(:, j)) < maxConnectionsPerColumn
                    matrix(i, j) = 1;  % Temporarily add the connection
                    [currentNODF, ~, ~] = calculate_NODF(matrix);  % Calculate NODF

                    if currentNODF > bestNODF  % Update if better NODF is found
                        bestNODF = currentNODF;
                        bestPosition = [i, j];
                    end

                    matrix(i, j) = 0;  % Revert the matrix to its original state
                end
            end
        end
    end

    % Break loop if no more connections can be added
    if isempty(bestPosition)
        break;
    end

    % Add the best connection found
    matrix(bestPosition(1), bestPosition(2)) = 1;
    % Compute Qb, Qr, NODF, and Num_mod
    [Qb, Qr, NODF_N, NODF_cols, NODF_rows, Num_mod] = computeQbAndMetrics(matrix);

    % Save the matrix for this step
    matrixFilename = ['9bus_bipartite_optimization_v2_', num2str(size(outputData, 1) + 1), '.xlsx'];
    writematrix(matrix, matrixFilename);
    disp(['Matrix saved to ', matrixFilename]);

    % Save the result for this step, including Num_mod
    outputData = [outputData; bestPosition(1), bestPosition(2), Qb, Qr, NODF_N, bestNODF, NODF_cols, NODF_rows, Num_mod];

    disp(['New connection added at position: (', num2str(bestPosition(1)), ', ', num2str(bestPosition(2)), ')']);
    disp(['Qb: ', num2str(Qb), ', Qr: ', num2str(Qr), ', N: ', num2str(NODF_N), ...
          ', NODF_cols: ', num2str(NODF_cols), ', NODF_rows: ', num2str(NODF_rows), ', Num_mod: ', num2str(Num_mod)]);
end


% Combine headers and data into a cell array
headers = {'Row', 'Column', 'Qb', 'Qr', 'NODF_N', 'bestNODF', 'NODF_cols', 'NODF_rows', 'Num_mod'};
combinedData = [headers; num2cell(outputData)];

% Save the combined data to an Excel file
resultsFilename = '9bus_bipartite_optimization_v2_.xlsx';
writecell(combinedData, resultsFilename);
disp(['Results saved to ', resultsFilename]);


% Function to calculate NODF and related metrics
function [NODF_N, NODF_cols, NODF_rows] = calculate_NODF(matrix)
    bp = Bipartite(matrix);
    nest_set = bp.nestedness.Detect();
    NODF_N = nest_set.N;
    NODF_cols = nest_set.N_cols;
    NODF_rows = nest_set.N_rows;
end

% Function to compute Qb, Qr, NODF, and Num_mod
function [Qb, Qr, NODF_N, NODF_cols, NODF_rows, Num_mod] = computeQbAndMetrics(matrix)
    bp = Bipartite(matrix);
    mod_set = bp.community.Detect();
    Num_mod = mod_set.N;  % Assign Num_mod
    Qb = mod_set.Qb;
    Qr = mod_set.Qr;
    [NODF_N, NODF_cols, NODF_rows] = calculate_NODF(matrix);
end
