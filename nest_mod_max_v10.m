% Add the bimat package to the path
addpath(genpath('/Users/emilypayne/Documents/bimat_packadge'));

% Read the bipartite matrix from the Excel file
matrix = readmatrix('tripartite matrix and results 9 bus.xlsx', 'Sheet', 'normal cps cyber', 'Range', 'B2:V22');

% Initialize variables
[m, n] = size(matrix);  
outputData = [];  % Store results for Excel

% Loop to add all possible connections one by one
while true
    bestNODF = -inf;
    bestPosition = [];

    % Find the best position to add a new connection
    for i = 1:m
        for j = 1:n
            if matrix(i, j) == 0
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

    % Break loop if no more connections can be added
    if isempty(bestPosition)
        break;
    end

    % Apply the optimal connection
    matrix(bestPosition(1), bestPosition(2)) = 1;

    % Compute Qb, Qr, NODF, and Num_mod
    [Qb, Qr, NODF_N, NODF_cols, NODF_rows, Num_mod] = computeQbAndMetrics(matrix);

    % Save the matrix for this step
    matrixFilename = ['optimized_matrix_step_', num2str(size(outputData, 1) + 1), '.xlsx'];
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
resultsFilename = 'connection_results.xlsx';
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

