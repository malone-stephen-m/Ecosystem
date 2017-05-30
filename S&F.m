function [] = Structure_and_Flow_FoodWebMetrics_calc_Laytonv_6_1
% Structural Based Food Web Metrics Calculations
% Written by Dr. Astrid Layton and Stephen Malone
% Copyright Georgia Institute of Technology
% Last Edit by Stephen Malone on Jan 27, 2016

% This code can be used to calculate structural based food web
%metrics for any network.
% The results of this code have been checked against the program enaR by
%Stuart Borrett and the work of Ulanowicz
%==========================================================================
%Important Notes

%If modeling one EIP with variations, user must specify changes in
%increasing complexity within the workbook with the final matrix being on 
%the last sheet in the Excel Workbook.

%If on a Linux/MAC machine, include path for JAR's transmitted with this script.
%Path for JXL and MXL JAR files for linux users
addpath('/home/ubuntubertserver/Documents/MATLAB/JARs/');

% An example of the formatting criteria for an EIP matrix is below:
%
    % | A   B    C                              D   E   F   G   H   I   J 
    %-|-------------------------------------------------------------------
    %1|                                          from						
    %2|    		actor                       	1	2	3	4	5	6	7
    %3| t	1	hay farm                    	0	1	0	0	0	0	0
    %4| o	2	wastewater treatment facility	0	0	1	0	0	0	0
    %5|    	3	pharmaceutical firms           	0	0	0	1	1	0	0
    %6|    	4	cogeneration facility       	0	0	1	0	0	0	0
    %7|    	5	waste management firms      	0	0	1	0	0	0	0
    %8|    	6	paint manufacture           	0	0	0	0	1	0	0
    %9|    	7	energy recovery             	0	0	0	0	1	0	0
%

%==========================================================================
% Program Operation:

% Once the program starts, you will be asked to open an excel file
%containing the data for which you would like to analyze. Once opened, the
%program will give you two options for 'Single EIP with variations or 
%'Multiple EIPs'. Select which method describes your analysis best. The 
%program will then begin updating the console with the number of sheets
%analyzed. If the workbook does not follow the correct format, an error
%will be returned. 
% The program will then generate images which are placed in the
%Generated_Images folder in the same directory which your original data set
%was analyzed from. 
% Once the program has completed its analysis, the calculated statistics
%will be located in the same folder from which the original was located
%with a prefix "Analyzed_". 
% The first sheet of the analysis labeled "stats" will contain all of the
%values from each sheet in one convenient location for easy comparison.
%Each sheet will also contain the original matrix with the stats listed
%below the matrix. 
%==========================================================================


% Allow user to open file anywhere on their computer
% uigetfile({'*.xlsx;*.xls' 'Excel file or Excel document';'*.xlsx' 'Excel file'; '*.xls' 'Excel document'},'Select a file','/home/ubuntubertserver/Documents/MATLAB');
        if ismac || isunix
            [filename, pathname] = uigetfile({'*.xlsx;*.xls' 'Excel file or Excel document';'*.xlsx' 'Excel file'; '*.xls' 'Excel document'},'Select a file','/home/ubuntubertserver/Documents/SustainableSteelProject');
        else
            [filename, pathname] = uigetfile({'*.xlsx;*.xls' 'Excel file or Excel document';'*.xlsx' 'Excel file'; '*.xls' 'Excel document'},'Select a file','C:\');
        end
        dt = questdlg('Is this workbook one EIP with variations, or multiple EIPs?','Question','One EIP','Multiple EIPs','Multiple EIPs');

        switch dt
            case 'One EIP'
                designType= '1';
            case 'Multiple EIPs' 
                designType = '2';
        end

        %Check for Cancel button press
        if (filename == 0)
            error('Input file is not selected!');
        end
        
        % Combine filename and pathname 
        fpfn = fullfile(pathname,filename);
        [~,name,~] = fileparts(filename);
        
        
        outFile = strcat('Analyzed_',name,'.xls');
        fpfno = fullfile(pathname,outFile);
        
        %Delete old analysis if one exists
        if exist(fpfno, 'file') == 2
            delete(fpfno);
        end
        
        %Read file info
        [~,sheets] = xlsfinfo(fpfn);
        numOfSheets = numel(sheets); 
        
        % Preallocate cell with the sheetcount for preformance.
        sheetdata = cell(numOfSheets,4);
        
        %Preallocate the info for each sheet which will be size of number
        %of sheets w/ each sheet showing matrix start point, size, and
        %whether it is a structure or flow matrix
        %[begIndexRow, begIndexCol,smatrixDim1, smatrixDim2, matrixType, actors, Smatrix,fmatrixDim1, fmatrixDim2, Fmatrix]
        startCell = cell(numOfSheets, 10);
        
        % Again preallocating for speed. We want to use the stored stats to
        %aggregate all the information we've gathered for later
        storedstats = cell(numOfSheets,11);

        sStatNamesc = {'Industrial_Park_Name','Cyclicity', 'Linkage_Density', 'Predator_Prey_Ratio', 'Generalization',...
                      'Vulnerability', 'Actors', 'Links', 'Number_of_Predators', 'Number_of_Prey','Connectance'};

        % Images Path and Printing
        imgdir = strcat(pathname,'Generated_Images');
        
        % Create the folder if it doesn't exist already.
        if ~exist(imgdir, 'dir')
          mkdir(imgdir);
        end

        % Clear Directory from past analysis
        if any(size(dir([imgdir '/*.jpg' ]),1))
            delete(strcat(imgdir,'*.jpg'));
        end
        
        
        %Pass each sheet from the workbook to findMatrix
        for i = 1:numOfSheets
            %Read in excel file
            [num,txt,raw] = xlsread(fpfn,i);
            
            status = sprintf('Reading Sheet %i',i);
            disp(status);
            [startCell{i,1},startCell{i,2},startCell{i,3},startCell{i,4},startCell{i,5}, startCell{i,6}, startCell{i,7},...
                startCell{i,8}, startCell{i,9}, startCell{i,10}]= findMatrix(raw, num, txt);
            
             %Name of Worksheet
            sheetdata{i,1} = sheets(1,i);
            
            %Actor Names
            sheetdata{i,2} = raw;
            sheetdata{i,3} = num;
            sheetdata{i,4} = txt;
        end
    warning('off','MATLAB:xlswrite:AddSheet');  
    %%
    %MULTIPLE EIPS
    if(designType == '2')
        % Loop through the sheets 
        for i = 1:numel(sheets)
            % If the data in in the right structure, sheets with the begenning
            %cells as NaN (Not a Number) values will be analyzed.
            %[begIndexRow, begIndexCol,smatrixDim1, smatrixDim2, matrixType, actors, Smatrix,fmatrixDim1, fmatrixDim2, Fmatrix]

            if startCell{i,5} ~= 'N'
                %Refine Actor names
                Actors = startCell{i,6}(:,:);

                %Different layout options are available and below. Pick which
                %one works best for you. For this study, equilibrium and
                %hierachical displayed best. For smaller matricies, equilibrium
                %worked best. For larger, hierachical was easier to view
                if startCell{i,5} == 'S'
                    % Refine matrix
                    Matrix = startCell{i,7};
                    bg = biograph(Matrix,Actors,'LayoutType','equilibrium','ShowTextInNodes','Label');
                else
                    Matrix = startCell{i,7};
                    %newActors  = ['imports'; Actors(:,1); 'fill1'; 'fill2'];
                    bg = biograph(Matrix,Actors,'LayoutType','equilibrium','ShowTextInNodes','Label');
                end
                %Option 2
                %bg = biograph(Matrix',Actors,'LayoutType','hierarchical','ShowTextInNodes','Label');
                %Option 3: 
                %bg = biograph(Matrix',Actors,'LayoutType','radial','ShowTextInNodes','Label');

                g = biograph.bggui(bg);
                f = get(g.biograph.hgAxes, 'Parent');

                % Print image to specified folder 
                imgname = strcat(imgdir, '/',sheetdata{i,1},'.jpg');
                try
                    print(f, '-djpeg', imgname);
                catch
                    print(f, '-djpeg', imgname{1})
                end
                
                % Close biograph GUI
                close(g.hgFigure);

                % Calculate Structure Metrics
                [num_Actors, num_Links, density_Link, connectance, num_Prey, num_Pred, num_SpecialPred,...
                    ratio_PreyToPred, fraction_SpecialPred, generalization,vulnerability, cyclicity, indGV] = Structure_Stats(Matrix);
                
                %If it is a flow matrix, calculate structure and flow
                if (startCell{i,5} == 'F')
                    [CI,MPL,AMI,ASC,DC,TSO,TST,alpha,R,H] = FlowBasedMetrics(startCell{i,10});
                    
                    %Collect Flow Stats
                    fStat = [CI,MPL,AMI,ASC,DC,TSO,TST,alpha,R,H];
                end

                % Had to add because cyclicity was returning a blank cell and
                %messing things up
                if isempty(cyclicity)
                    cyclicity =0;
                end

                % Create labels for structure statistics 
                sStatNames = {'Cyclicity', 'Linkage Density', 'Predator/Prey Ratio', 'Generalization',...
                'Vulnerability', 'Actors', 'Links', 'Number of Predators', 'Number of Prey',...
                'Connectance', 'Number of Special Prey', 'Fraction of Special Predators'};
            
                %Create Labels for flow stats
                fStatNames = {'Finn Cycling Index','Mean Path Length','Average Mutual Information','Ascendency',...
                    'Developmental Capacity','Total System Overhead','Total System Through Flow','Alpha', 'Robustness', 'Shannon Index'};

                % Collect statistics in order of labels
                sStat = [cyclicity, density_Link, ratio_PreyToPred, generalization, vulnerability, ...
                    num_Actors, num_Links, num_Pred, num_Prey, connectance, num_SpecialPred, fraction_SpecialPred];
                
                if startCell{i,5} == 'S'
                    % Combine labels and statistics for easy viewing
                    sStats1 = padconcatenation(sStat,indGV,1);
                    CombineStats = {sStatNames, num2cell(sStats1)};
                    %Find Maximum Column Count of Stats
                    [maxSize,IndOfMax] = max(cellfun('length',CombineStats),[],2);
                    [~, IndOfMaxCol] = max(sum(cellfun(@(c) size(c,1),CombineStats),1));

                    padded_CombinedStats = cell(1,numel(CombineStats));
                    %Pad Names and Stats in Order to Concatenate
                    for jj = 1:numel(CombineStats)
                        if jj == IndOfMax
                            padded_CombinedStats{jj} = CombineStats{jj}; 
                            continue
                        else
                            padded_CombinedStats{jj} = csr_pad(CombineStats{jj},IndOfMaxCol, IndOfMax, maxSize, CombineStats{jj}{1,1});
                        end
                    end
                    sStats = [padded_CombinedStats{1}; padded_CombinedStats{2}];                    
                else
                    %Combining Flow stats and labels with structure 
                    sStats1 = padconcatenation(fStat,indGV,1);

                    %Combine Stats to Write
                    CombineStats = {sStatNames, num2cell(sStat), fStatNames, num2cell(sStats1)};
                    %Find Maximum Column Count of Stats
                    [maxSize,IndOfMax] = max(cellfun('length',CombineStats),[],2);
                    [~, IndOfMaxCol] = max(sum(cellfun(@(c) size(c,1),CombineStats),1));
                    padded_CombinedStats = cell(1,numel(CombineStats));
                    %Pad Names and Stats in Order to Concatenate
                    for jj = 1:numel(CombineStats)
                        if jj == IndOfMax
                            padded_CombinedStats{jj} = CombineStats{jj}; 
                            continue
                        else
                            %csr_pad(cell_in, Ind1, Ind2, max_length, char)
                            padded_CombinedStats{jj} = csr_pad(CombineStats{jj},IndOfMaxCol, IndOfMax, maxSize, CombineStats{jj}{1,1});
                        end
                    end
                    sStats = [padded_CombinedStats{1};padded_CombinedStats{2}; padded_CombinedStats{3}; padded_CombinedStats{4}];
                end
                

                %Find location to write to on the sheet. 
                [temp_sheetdata,temp_sStat] = format_RemoveStrings(sheetdata{i,2},sStats); 
                C = padconcatenation(temp_sheetdata,temp_sStat,1);
                C = num2cell(C);
                [s1,s2] = size(sheetdata{i,2});
                C(1:s1, 1:s2) = sheetdata{i,2};
                if startCell{i,5} == 'S'
                    filler2 = csr_pad(sStatNames, 1, 1, length(C(s1+3,1:end)), sStatNames{1,1});
                    C(s1+1,1:end) = filler2;
                elseif startCell{i,5} == 'F'
                    filler2 = csr_pad(sStatNames, 1, 1, length(C(s1+3,1:end)), sStatNames{1,1});
                    C(s1+1,1:end) = filler2;
                    filler = csr_pad(fStatNames, 1, 1, length(C(s1+3,1:end)), fStatNames{1,1});
                    C(s1+3,1:end) = filler;
                end
                
                    
                %Replace NaNs with Blanks
                for k = 1:numel(C)
                  if isnan(C{k})
                    C{k} = '';
                  end
                end
                                    
                robust_Write(fpfno,C,'A1', sheetdata{i,1}, 'results');

                % Used to check Stats on Astrids Thesis
                storedstats{i,1} = sheetdata{i,1};
                storedstats{i,2} = cyclicity;
                storedstats{i,3} = density_Link;
                storedstats{i,4} = ratio_PreyToPred;
                storedstats{i,5} = generalization;
                storedstats{i,6} = vulnerability;
                storedstats{i,7} = num_Actors;
                storedstats{i,8} = num_Links;
                storedstats{i,9} = num_Pred;
                storedstats{i,10} = num_Prey;
                storedstats{i,11} = connectance;

                % Update the user on progress
                message = sprintf('Writing Sheet: %i',i);
                disp(message);
                
                % If the end of our file is reached
                if i == numel(sheets)

                    % Remove any empty rows from the cell containing our stats
                    storedstats(all(cellfun('isempty',storedstats),2),:) = [];

                    % Create Table with our stats
                    statsTable = cell2table(storedstats, 'VariableNames', sStatNamesc);

                    % Sort the table by most weighted variable
                    [sortedStatsTable,~] = sortrows(statsTable,{'Cyclicity','Linkage_Density', 'Predator_Prey_Ratio',...
                        'Generalization','Vulnerability'},{'descend','descend','descend','descend','descend'});

                    % Concatenate column headers with data
                    finalStats = [sortedStatsTable.Properties.VariableNames;table2cell(sortedStatsTable)];              

                    % Write the aggregated stats if the file is done
%                    xlswrite(fpfno, finalStats, 'Stats');
                    robust_Write(fpfno,finalStats,'A1', 'Stats', 'Stats');
                
                end
                
            else
                if (isempty(txt))
                    continue;
                end
                
                robust_Write(fpfno,sheetdata{i,4},'A1', sheets{1,i}, 'Sheets');

                % Let the user know the progress
                message = sprintf('Writing Sheet: %i',i);
                disp(message);

                % If the end of our file is reached
                if i == numOfSheets

                    % Remove any empty rows from the cell containing our stats
                    storedstats(all(cellfun('isempty',storedstats),2),:) = [];

                    % Create Table with our stats
                    statsTable = cell2table(storedstats, 'VariableNames', sStatNamesc);

                    % Sort the table by most weighted variable
                    [sortedStatsTable,~] = sortrows(statsTable,{'Cyclicity','Linkage_Density', 'Predator_Prey_Ratio',...
                        'Generalization','Vulnerability'},{'descend','descend','descend','descend','descend'});

                    % Concatenate column headers with data
                    finalStats = [sortedStatsTable.Properties.VariableNames;table2cell(sortedStatsTable)];              

                    robust_Write(fpfno,finalStats,'A1', 'Stats', 'Stats');
                end
                continue;
            end
        end
        %%
        %ONE EIP
    elseif (designType == '1')
        double parentMatrix;
        first_Iteration = 1;
        for i = numel(sheets):-1:1
                % If the data in in the right structure, sheets with the begenning
                %cells as NaN (Not a Number) values will be analyzed.
                %[begIndexRow, begIndexCol,smatrixDim1, smatrixDim2, matrixType, actors, Smatrix,fmatrixDim1, fmatrixDim2, Fmatrix]               
                if startCell{i,5} ~= 'N'
                    %Refine Actor names
                    Actors = startCell{i,6}(:,:);

                    % Refine matrix
                    Matrix = startCell{i,7};

                    %Different layout options are available and below. Pick which
                    %one works best for you. For this study, equilibrium and
                    %hierachical displayed best. For smaller matricies, equilibrium
                    %worked best. For larger, hierachical was easier to view
                    if startCell{i,5} == 'S'
                        bg = biograph(Matrix,Actors,'LayoutType','equilibrium','ShowTextInNodes','Label');
                    else
                        Matrix = startCell{i,7};
                        bg = biograph(Matrix,Actors,'LayoutType','equilibrium','ShowTextInNodes','Label');
                    end
                    %Option 2
                    %bg = biograph(Matrix',Actors,'LayoutType','hierarchical','ShowTextInNodes','Label');
                    %Option 3: 
                    %bg = biograph(Matrix',Actors,'LayoutType','radial','ShowTextInNodes','Label');

                    %Get Parent Matrix Positions
                    if (i == numel(sheets))
                        dolayout(bg);
                        parentMatrix = get(bg.Nodes,'Position');
                    else
                        if first_Iteration == 1
                            try
                                dolayout(bg);
                                parentMatrix = get(bg.Nodes,'Position');
                            catch
                                if startCell{i-1,5} == 'S'
                                    bg = biograph(Matrix,Actors,'LayoutType','equilibrium','ShowTextInNodes','Label');
                                    dolayout(bg);
                                    parentMatrix = get(bg.Nodes,'Position');
                                else
                                    Matrix = startCell{i-1,7};
                                    bg = biograph(Matrix,Actors,'LayoutType','equilibrium','ShowTextInNodes','Label');
                                    dolayout(bg);
                                    parentMatrix = get(bg.Nodes,'Position');
                                end
                            end
                            first_Iteration = 0;
                       else
                            if(exist('parentMatrix','var'))
                                dolayout(bg);
                                %Set node locations to parent's
                                for z = 1:numel(Actors)
                                    bg.nodes(z).Position = cell2mat(parentMatrix(z));
                                end
                            else
                                error('Error. \n Please move any information sheets at end of workbook to the front of workbook for processing. We cannot find Data!')
                            end
                            dolayout(bg, 'Pathsonly', true);
                       end
                    end

                    g = biograph.bggui(bg);
                    f = get(g.biograph.hgAxes, 'Parent');

                    % Print image to specified folder 
                    imgname = strcat(imgdir, '/',sheetdata{i,1},'.jpg');
                    try
                        print(f, '-djpeg', imgname);
                    catch
                        print(f, '-djpeg', imgname{1})
                    end

                    % Close biograph GUI
                    close(g.hgFigure);

                    % Calculate Structure Metrics
                    [num_Actors, num_Links, density_Link, connectance, num_Prey, num_Pred, num_SpecialPred,...
                        ratio_PreyToPred, fraction_SpecialPred, generalization,vulnerability, cyclicity, indGV] = Structure_Stats(Matrix);

                    %If it is a flow matrix, calculate structure and flow
                    if (startCell{i,5} == 'F')
                        [CI,MPL,AMI,ASC,DC,TSO,TST,alpha,R,H] = FlowBasedMetrics(startCell{i,10});

                        %Collect Flow Stats
                        fStat = [CI,MPL,AMI,ASC,DC,TSO,TST,alpha,R,H];
                    end

                    % Had to add because cyclicity was returning a blank cell and
                    %messing things up
                    if isempty(cyclicity)
                        cyclicity =0;
                    end

                    % Create labels for structure statistics 
                    sStatNames = {'Cyclicity', 'Linkage Density', 'Predator/Prey Ratio', 'Generalization',...
                    'Vulnerability', 'Actors', 'Links', 'Number of Predators', 'Number of Prey',...
                    'Connectance', 'Number of Special Prey', 'Fraction of Special Predators'};

                    %Create Labels for flow stats
                    fStatNames = {'Finn Cycling Index','Mean Path Length','Average Mutual Information','Ascendency',...
                        'Developmental Capacity','Total System Overhead','Total System Through Flow','Alpha', 'Robustness', 'Shannon Index'};

                    % Collect statistics in order of labels
                    sStat = [cyclicity, density_Link, ratio_PreyToPred, generalization, vulnerability, ...
                        num_Actors, num_Links, num_Pred, num_Prey, connectance, num_SpecialPred, fraction_SpecialPred];

                    if startCell{i,5} == 'S'
                        % Combine labels and statistics for easy viewing
                        sStats1 = padconcatenation(sStat,indGV,1);
                        CombineStats = {sStatNames, num2cell(sStats1)};
                        %Find Maximum Column Count of Stats
                        [maxSize,IndOfMax] = max(cellfun('length',CombineStats),[],2);
                        [~, IndOfMaxCol] = max(sum(cellfun(@(c) size(c,1),CombineStats),1));

                        padded_CombinedStats = cell(1,numel(CombineStats));
                        %Pad Names and Stats in Order to Concatenate
                        for jj = 1:numel(CombineStats)
                            if jj == IndOfMax
                                padded_CombinedStats{jj} = CombineStats{jj}; 
                                continue
                            else
                                padded_CombinedStats{jj} = csr_pad(CombineStats{jj},IndOfMaxCol, IndOfMax, maxSize, CombineStats{jj}{1,1});
                            end
                        end
                        sStats = [padded_CombinedStats{1}; padded_CombinedStats{2}];                    
                    else
                        %Combining Flow stats and labels with structure 
                        sStats1 = padconcatenation(fStat,indGV,1);

                        %Combine Stats to Write
                        CombineStats = {sStatNames, num2cell(sStat), fStatNames, num2cell(sStats1)};
                        %Find Maximum Column Count of Stats
                        [maxSize,IndOfMax] = max(cellfun('length',CombineStats),[],2);
                        [~, IndOfMaxCol] = max(sum(cellfun(@(c) size(c,1),CombineStats),1));
                        padded_CombinedStats = cell(1,numel(CombineStats));
                        %Pad Names and Stats in Order to Concatenate
                        for jj = 1:numel(CombineStats)
                            if jj == IndOfMax
                                padded_CombinedStats{jj} = CombineStats{jj}; 
                                continue
                            else
                                %csr_pad(cell_in, Ind1, Ind2, max_length, char)
                                padded_CombinedStats{jj} = csr_pad(CombineStats{jj},IndOfMaxCol, IndOfMax, maxSize, CombineStats{jj}{1,1});
                            end
                        end
                        sStats = [padded_CombinedStats{1};padded_CombinedStats{2}; padded_CombinedStats{3}; padded_CombinedStats{4}];
                    end


                    %Find location to write to on the sheet. 
                    [temp_sheetdata,temp_sStat] = format_RemoveStrings(sheetdata{i,2},sStats); 
                    C = padconcatenation(temp_sheetdata,temp_sStat,1);
                    C = num2cell(C);
                    [s1,s2] = size(sheetdata{i,2});
                    C(1:s1, 1:s2) = sheetdata{i,2};
                    if startCell{i,5} == 'S'
                        filler2 = csr_pad(sStatNames, 1, 1, length(C(s1+3,1:end)), sStatNames{1,1});
                        C(s1+1,1:end) = filler2;
                    elseif startCell{i,5} == 'F'
                        filler2 = csr_pad(sStatNames, 1, 1, length(C(s1+3,1:end)), sStatNames{1,1});
                        C(s1+1,1:end) = filler2;
                        filler = csr_pad(fStatNames, 1, 1, length(C(s1+3,1:end)), fStatNames{1,1});
                        C(s1+3,1:end) = filler;
                    end
                    
                    %Replace NaNs with Blanks
                    for k = 1:numel(C)
                      if isnan(C{k})
                        C{k} = '';
                      end
                    end
                    
                    
                    robust_Write(fpfno,C,'A1', sheetdata{i,1}, 'results');

                    % Used to check Stats on Astrids Thesis
                    storedstats{i,1} = sheetdata{i,1};
                    storedstats{i,2} = cyclicity;
                    storedstats{i,3} = density_Link;
                    storedstats{i,4} = ratio_PreyToPred;
                    storedstats{i,5} = generalization;
                    storedstats{i,6} = vulnerability;
                    storedstats{i,7} = num_Actors;
                    storedstats{i,8} = num_Links;
                    storedstats{i,9} = num_Pred;
                    storedstats{i,10} = num_Prey;
                    storedstats{i,11} = connectance;

                    % Update the user on progress
                    message = sprintf('Writing Sheet: %i',i);
                    disp(message);

                    % If the end of our file is reached
                    if i == 1

                        % Remove any empty rows from the cell containing our stats
                        storedstats(all(cellfun('isempty',storedstats),2),:) = [];

                        % Create Table with our stats
                        statsTable = cell2table(storedstats, 'VariableNames', sStatNamesc);

                        % Sort the table by most weighted variable
                        [sortedStatsTable,~] = sortrows(statsTable,{'Cyclicity','Linkage_Density', 'Predator_Prey_Ratio',...
                            'Generalization','Vulnerability'},{'descend','descend','descend','descend','descend'});

                        % Concatenate column headers with data
                        finalStats = [sortedStatsTable.Properties.VariableNames;table2cell(sortedStatsTable)];              

                        % Write the aggregated stats if the file is done
    %                    xlswrite(fpfno, finalStats, 'Stats');
                        robust_Write(fpfno,finalStats,'A1', 'Stats', 'Stats');
                    end

                else
                    if (isempty(txt))
                        continue;
                    end
                    robust_Write(fpfno,sheetdata{i,4},'A1', sheets{1,i}, 'Sheets');

                    % Let the user know the progress
                    message = sprintf('Writing Sheet: %i',i);
                    disp(message);

                    % If the end of our file is reached
                    if i == numOfSheets
                        if numOfSheets == 1
                            % Remove any empty rows from the cell containing our stats
                            storedstats(all(cellfun('isempty',storedstats),2),:) = [];

                            % Create Table with our stats
                            statsTable = cell2table(storedstats, 'VariableNames', sStatNamesc);

                            % Sort the table by most weighted variable
                            [sortedStatsTable,~] = sortrows(statsTable,{'Cyclicity','Linkage_Density', 'Predator_Prey_Ratio',...
                                'Generalization','Vulnerability'},{'descend','descend','descend','descend','descend'});

                            % Concatenate column headers with data
                            finalStats = [sortedStatsTable.Properties.VariableNames;table2cell(sortedStatsTable)];              

                            % Write the aggregated stats if the file is done
                            %xlswrite(fpfno, finalStats, 'Stats', 'A1');
                            robust_Write(fpfno,finalStats,'A1', 'Stats', 'Stats');
                        end
                    elseif i == 1
                            % Remove any empty rows from the cell containing our stats
                            storedstats(all(cellfun('isempty',storedstats),2),:) = [];

                            % Create Table with our stats
                            statsTable = cell2table(storedstats, 'VariableNames', sStatNamesc);

                            % Sort the table by most weighted variable
                            [sortedStatsTable,~] = sortrows(statsTable,{'Cyclicity','Linkage_Density', 'Predator_Prey_Ratio',...
                                'Generalization','Vulnerability'},{'descend','descend','descend','descend','descend'});

                            % Concatenate column headers with data
                            finalStats = [sortedStatsTable.Properties.VariableNames;table2cell(sortedStatsTable)];              

                            % Write the aggregated stats if the file is done
                            %xlswrite(fpfno, finalStats, 'Stats', 'A1');
                            robust_Write(fpfno,finalStats,'A1', 'Stats', 'Stats');
                    end
                
                    continue;
                end
        end
    end
   % Let user know the progress
   disp('Finished!');
   %%
end

function new_array = csr_pad(cell_in, Ind1, Ind2, max_length, char)
    %If input array is an array of Strings
    if ischar(char)
        %Pad end of array with a Blank string to get size right to concat
        cell_in(length(cell_in) + 1 : max_length) = {''};
        new_array =  cell_in;
    %Numerical Array
    elseif isnumeric(char)
        if (Ind1 == Ind2)
            %Pad end of array with NaN to get size right to concatenate
            new_array = num2cell(padarray([cell_in{1,:}],[0,max_length-length(cell_in)],NaN,'post'));
        else
            [numRows,~] = size(cell_in);
            new_array = cell(numRows,max_length);
            for ii = 1:numRows
                new_array(ii,:) = num2cell(padarray([cell_in{ii,:}],[0,max_length-length(cell_in)],NaN,'post'));
            end
        end
            
    end
end

function robust_Write(out_fn,txt2write,cell2write, sheet2write, type)

    jxlPath = '/home/ubuntubertserver/Documents/MATLAB/JARs/jxl.jar';
    mxlPath = '/home/ubuntubertserver/Documents/MATLAB/JARs/MXL.jar';
    if strcmp(type, 'Stats')
        if ispc
            % Write info that is not analyzed to the file
            xlswrite(out_fn,txt2write, sheet2write, cell2write);
        else
            try
                javaaddpath(jxlPath);
                javaaddpath(mxlPath);
                import mymxl.*;
                import jxl.*;  
                xlwrite(out_fn,txt2write,sheet2write);
            catch
                char_inds = cellfun(@ischar,txt2write);
                char_b = txt2write(char_inds);
                for t = 1:numel(char_b)
                    xlswrite(out_fn,txt2write, sheet2write{1,1}, cell2write);
                end
            end 
        end
    elseif strcmp(type, 'Sheets')
                if ispc
                    % Write info that is not analyzed to the file
                    xlswrite(out_fn,txt2write,sheet2write{1,1}, cell2write);
                else
                    try
                        javaaddpath(jxlPath);
                        javaaddpath(mxlPath);
                        import mymxl.*;
                        import jxl.*;  
                        xlwrite(out_fn,txt2write,sheet2write);
                    catch
                        char_inds = cellfun(@ischar,txt2write);
                        char_b = txt2write(char_inds);
                        for t = 1:numel(char_b)
                            cell2write = sprintf('A%i',t);
                            xlswrite(out_fn, char_b{t}, sheet2write, cell2write);
                        end
                    end 
                end
            
    elseif strcmp(type, 'results')
        if ispc
            %Windows Users

            % Write to file using the same format
            % Raw supplied data first
            xlswrite(out_fn,txt2write,sheet2write, cell2write);
        else
            try
                javaaddpath(jxlPath);
                javaaddpath(mxlPath);
                import mymxl.*;
                import jxl.*;  
                xlwrite(out_fn,txt2write,sheet2write{1,1});

            catch
                char_inds = cellfun(@ischar,txt2write);
                char_b = txt2write(char_inds);
                for t = 1:numel(char_b)
                    cell2write = sprintf('A%i',t);

                    xlswrite(out_fn, char_b{t}, sheet2write{1,1}, cell2write);
                    
                end
            end 
        end
    end
end

function [strRem1, strRem2] = format_RemoveStrings(input1,input2)
    input1(cellfun(@(x) ischar(x),input1)) = {NaN};
    input2(cellfun(@(x) ischar(x),input2)) = {NaN};
    strRem1 = cell2mat(input1);
    strRem2 = cell2mat(input2);
end
        
function [begIndexRow, begIndexCol,smatrixDim1, smatrixDim2, matrixType, ...
    actors, Smatrix,fmatrixDim1, fmatrixDim2, Fmatrix] = findMatrix(rawWorkSheet, numWorkSheet, txtWorkSheet)
%Determine Where Matrix Begins in Spreadsheet
%   First find where the matrix begins, then determine size. Finally,
%   determine if the matrix is a structure or flow matrix. ROWS MUST BE
%   NUMBERED HORIZONTALLY FOR THIS FUNCTION TO WORK CORRECTLY

    f = rawWorkSheet;
    g = rawWorkSheet;
    h = txtWorkSheet;
    f(cellfun(@ischar,f))={NaN};
    out = cell2mat(f);
    
    %replace all Nans
    rawWorkSheet(cellfun(@(X) any(isnan(X)),rawWorkSheet)) = {''};
    coi = cellfun(@isnumeric, rawWorkSheet);

    %if there are no cells of interest, return original cells and alert to data
    %being unusable
    b = any(coi(:) > 0);
    if b > 0
%         
%         %Determine index of coi within the raw file
%         idxs = cellfun(@(x)find(coi==x,1),rawWorkSheet);

        %calculate size of numbers in spreadsheet
        [s1,~] = size(numWorkSheet);
         
         %Change all NaN to zero for isSorted Analysis
         numWorkSheet(isnan(numWorkSheet)) = 0;
         
         %Check to see if this is the first line in the Matrix
         isFirstLine = 1;
         
         %Loop through rows to see if they are just the numbers to count
         %actors and remove them if so. Exclude rows that are all zero.Must
         %handle rows after first counting line that could be sorted and
         %non-zero as well. 
         for t = 1:s1
             if issorted(numWorkSheet(t,1:length(numWorkSheet)-3)) && any(numWorkSheet(t,:))
                 if isFirstLine == 1
                     %Find row and column which counting begins
                     [row, col] = find(numWorkSheet(t,:) == 1);
                     %If any values contained in the matrix exceed 1
                      if sum(any(numWorkSheet(row+1:end, col:end)>1))
                         numWorkSheet(1:t,:) = [];
                         Fmatrix = numWorkSheet(row:end,col-1:end);
                         %matrix = numWorkSheet(:,2:end);
                         matrixType = 'F';
                         %Determine dimensions
                         [fmatrixDim1, fmatrixDim2] = size(Fmatrix);
                         %Return Location
                         [begIndexRow,begIndexCol]= findsubmat(out,Fmatrix);
                         %Return Actors
                         actors = g(begIndexRow+1:begIndexRow+fmatrixDim1-3,begIndexCol-1);
                         
                         %SMAtrix
                         T = Fmatrix;
                         %Reformat Matrix for Structure Stats
                          [numRow, numCol] = size(T);
                          T(T~=0) = 1;
                          %remove First Row and column
                          T(1,:) =[];
                          T(:,1) =[];
                          %Remove Last Two Rows and columns
                          T(numRow-2:numRow-1,:) = [];
                          T(:,numCol-2:numCol-1) = [];
                          [smatrixDim1, smatrixDim2] = size(T);
                          Smatrix = T;
                         return;
                      else
                         %Delete row w/ counting
                         numWorkSheet(1:t,:) = [];
                         Smatrix = numWorkSheet(row:end,col:end);
                         %return Structure
                         matrixType = 'S';
                         %Determine dimensions
                         [smatrixDim1, smatrixDim2] = size(Smatrix);
                         %return Location
                         [begIndexRow,begIndexCol]= findsubmat(out,Smatrix);
                         Fmatrix =0; 
                         fmatrixDim1=0; 
                         fmatrixDim2 = 0;
                        %return Actors
                        actors = g(begIndexRow:begIndexRow+smatrixDim1-1,begIndexCol-1);
                        return;
                      end

                 else
                     break;
                 end           
             end
             if t == s1
                 break;
             end
         end
    else
        %return data on sheet without changes
        matrixType = 'N';
        Smatrix = 0;
        begIndexRow = 0;
        begIndexCol = 0;
        smatrixDim1 = 0;
        smatrixDim2 = 0;
        Fmatrix = 0;
        fmatrixDim1 = 0;
        fmatrixDim2 = 0;
        actors = h;
    end
end

function [idx,idx2] = findsubmat(A,B)
% FINDSUBMAT find one matrix (a submatrix) inside another.
% IDX = FINDSUBMAT(A,B) looks for and returns the linear index of the
% location of matrix B within matrix A.  The index IDX corresponds to the
% location of the first element of matrix B within matrix A.
% [R,C] = FINDSUBMAT(A,B) returns the row and column instead.
%
% EXAMPLES:
%
%         A = magic(12);
%         B = [81 63;52 94];
%         [r,c]= findsubmat(A,B)
%         % Now to verify:
%         A(r:(r+size(B,1)-1),c:(c+size(B,2)-1))==B
%
%         A = [Inf NaN -Inf 3 4; 2 3 NaN 6 7; 5 6 3 1 6];
%         B = [Inf NaN -Inf;2 3 NaN];
%         findsubmat(single(A),single(B))
%
% Note:  The interested or concerned user who wants to know how NaNs are 
% handled should see the extensive comment block in the code.  
%
% See also find, findstr, strfind
%
% Author:   Matt Fig with improvements by others (See comments)
% Contact:  popkenai@yahoo.com
% Date:     3/31/2009
% Updated:  5/5/2209 to allow for NaNs in B, as per Urs S. suggestion. 
%                    Also to return [] if B is empty.
%           6/12/2009 Introduce another algorithm for larger B matrices.

if nargin < 2
    error('Two inputs are required.  See help.')
end

[rA,cA] = size(A);  % Get the sizes of the inputs. First the larger matrix.
[rB,cB] = size(B);  % The smaller matrix.
tflag = false; % In case we need to transpose.

if ~ismatrix(A) || ~ismatrix(B) || rA<rB || cA<cB
    error('Only 2D arrays are allowed.  B must be smaller than A.')
elseif isempty(B)
    idx = [];  % This should be an obvious situation.
    idx2 = [];
    return;
elseif rA==rB && cA==cB  % User is wasting time, annoy him with a disp().
    disp(['FINDSUBMAT is not recommended for equality determination, ',...
          'instead use:  isequal(A,B)']);
    idx = [];
    idx2 = [];
    
    if all(A(:)==B(:))
        idx = 1;
        idx2 = 1;
    end
    return;
elseif rB==1 && cB==1  % User is wasting time, annoy him with a disp().
    disp(['FINDSUBMAT is not recommended for finding scalars, ',...
          'instead use:  find(A==B)']);
    if nargout==2
        [idx,idx2] = find(A==B);
    else
        idx = find(A==B);
    end
    return;
end

if cB > ceil(1.5*rB)
    A = A.';  % In this case it may be faster to transpose.
    B = B.';  % The 1.5 cutoff is based on several trial runs.
    [rA,cA] = size(A);  % Get the sizes of the inputs transposed.
    [rB,cB] = size(B);
    tflag = true; % For the correct output at the end.
end

nans = isnan(B(:));  % If B has NaNs, user expects to find match in A.

if any(nans)
    % There are at least two strategies for dealing with NaNs here.  One 
    % approach is to pick the largest finite number N between B and A,
    % then replace the NaNs in both matrices with (N + 1).  This has the
    % advantage of certainty when it comes to uniqueness.  Unfortunately,
    % this is slow for large problems.  The other approach is to assign the
    % NaNs in A and B to an 'unlikely' number.  This is much faster, but 
    % also has the risk of duplicating other elements in A or B and thereby
    % giving false results.  The odds against a conflict can be minimized
    % by choosing the unlikely number with some care.  First, the number
    % should not be on [0 1] because this is a very common range, often
    % encountered when working with: images, matrices derived using rand,
    % and normalized data.  Second, the number should not be an integer or
    % an explicit ratio of integers (3/5, 45673/2344), for obvious reasons.
    % Third, in this case we want the number to be a function of rand.  The
    % reason is that for very large problems, the first method above takes
    % up to 20+ times longer than the unlikely number method.  Thus if a
    % user is paranoid about trusting the output when using this method, 
    % even with all of the above precautions and exclusions, it will 
    % still be much faster to run the code twice and compare answers than 
    % to run the code once using the first method, with a few exceptions.
    % I have chosen speed over certainty, but also include the alternate 
    % method for convenience.  To switch to the first approach, uncomment 
    % the first line and comment out the second line below.
%     vct = max([B(isfinite(B(:)))',Atv(isfinite(Atv))]) + 1; % Certainty.
    vct = pi^pi -1/pi + 1/rand; % Unlikely number on [37.14...   9.0...e15]
    B(nans) = vct;  % Set the NaNs in both to vct.
    A(isnan(A)) = vct;
    clear nans  % This could be a large vector, save some memory.
end

if numel(B)<30  % The below method is faster for most small B matrices.
    A = A(:).';  % Make a single row vector for strfind
    vct = strfind(A,B(:,1).');  % Find occurrences of the first col of B.
    % Next eliminate wrap-arounds, this was much improved by Bruno Luong.
    idx = vct(mod(vct-1,rA) <= rA-rB);
    cnt = 2;  % Counter for the while loop.

    while ~isempty(idx) && cnt<=cB
        vct = strfind(A,B(:,cnt).');  % Look for successive columns.
        % The C helper function ismembc needs both args to be sorted.
        % Search the code in ismember for more information.
        idx = idx(ismembc(idx + (cnt-1)*rA,vct)); % Matches with previous?
        cnt = cnt+1;  % Increment counter.
    end
else % The below method is faster for most larger B matrices.
    idx = strfind(A(:).',B(:,1).');  % Occurrences of the first col of B.
    % Next eliminate wrap-arounds, this was much improved by Bruno Luong.
    idx = idx(mod(idx-1,rA) <= rA-rB);
    idx(idx>((cA-cB+1)*rA)) = []; % Too close to right edge.
    cnt = 2;  % Counter for the while loop.
    flag = true(1,length(idx));
    % Siyi Deng noticed that for large B, the previous algorithm was slow.
    % The code below reflects an account for this behavior.
    while cnt<=cB
        TMP = rA*(cnt-1);  % Just to make the cond below more intelligible.
        TMP2 = B(:,cnt)';
        for jj = 1:length(idx)
            if ~isequal(A(idx(jj)+TMP : idx(jj)+rB-1+TMP),TMP2)
                flag(jj) = false;
            end
        end
        cnt = cnt+1;  % Increment counter.
    end

    idx = idx(flag);
end

if tflag  % We must get the index to A from index to A'
    tmp = rem(idx-1,rA) + 1;  % A temporary variable.
    idx = 1 + (idx - tmp)/rA + (tmp - 1)*cA;
    rA = cA;  % Only used if user wants subscripts instead of Linear Index.
end

if nargout==2 % Get subscripts.
    tmp = rem(idx-1,rA) + 1; 
    idx2 = (idx - tmp)/rA + 1;
    idx = tmp;                        
end
end

%Structure Metrics
function [num_Actors, num_Links, density_Link, connectance, num_Prey, num_Pred,...
    num_SpecialPred, ratio_PreyToPred, fraction_SpecialPred, generalization,vulnerability, cyclicity, indGV] = Structure_Stats( input_Matrix )
%========================== Structural Based Metrics ======================
    com = input_Matrix;
    %Standard metrics
    
    %number of actors in network
    num_Actors = size(input_Matrix,1);            
    %calculating number of links
    num_Links = sum(sum(com));      	
    %link density
    density_Link = num_Links/num_Actors;               
    %connectance
    connectance = num_Links/num_Actors^2;        
    
    %Predator and Prey metrics

    %number of predators for each prey
    com_row_sum = sum(com,2);        
    %number of prey for each predator
    com_col_sum = sum(com);   

    %number of prey
    num_Prey = nnz(com_row_sum);             
    %number or predators
    num_Pred = nnz(com_col_sum); 


    %determines number of predators with only one prey
    num_SpecialPred = sum(eq(com_col_sum,1));  
    %prey to predator ratio
    ratio_PreyToPred = num_Prey/num_Pred;                 
    %fraction specialized predators
    fraction_SpecialPred = num_SpecialPred/num_Pred;  
    
    indGV = zeros(3,num_Actors);
    for i = 1:length(com_col_sum)
        indGV(1,i) = i;
        %Individual Generalization
        indGV(2,i) = sum(com_col_sum(i))/num_Pred;
        %Individual Vulnerability
        indGV(3,i) = sum(com_row_sum(i))/num_Prey;
    end
    
    generalization = sum(com_col_sum)/num_Pred;
    vulnerability = sum(com_row_sum)/num_Prey;
    %Advanced Structural metrics for networks
    eigenvalues = eig(com);       
    %cyclicity, the max. eigenvalue
    cyclicity = max(eigenvalues(abs(real(eigenvalues))== abs(eigenvalues)));
end

%Flow Metrics
function [CI,MPL,AMI,ASC,DC,TSO,TST,alpha,R,H] = FlowBasedMetrics(Matrix)
%Flow Based Metrics Calculations
% 'Finn Cycling Index','Mean Path Length','Average Mutual Information','Ascendency','Developmental Capacity','Total System Overhead','Alpha', 'Robustness', 'Shannon Index'

%========================== Flow Based Metrics ============================
% FLOW MATRIX
% Enter flow matrix [T] below and uncomment
% Must be a square matrix of size (N+3) x (N+3) where N represents the number of actors in the system
% Flow is documented as moving from row (i) to column (j). The entry in row i and column j would be denoted as tij.
% Row 0 represents imports from outside the system
% Rows 1 through N represent exchanges inside the system between actors
% Columns N+1 and N+2 represent exports to outside the system. Column N+1 is usable exports (still have value) and Column N+2 is unusable (all value is gone)
% Rows N+1 and N+2 should be empty
% Column 0 should be empty
% for example: T = [ 0	288.6	806.4	0	0	0 ;...
%                    0	0	557.7	0	0	0 ;...
%                    0	0	0	1364	0	0 ;...
%                    0	269.1	0	0	1530.2	0 ;...
%                    0	0	0	0	0	0 ;...
%                    0	0	0	0	0	0 ];


T = Matrix;  % Enter flow matrix [T] here and UNCOMMENT this line

n_T = size(T,1);
        T_csum = sum(T,1);  %sum over columns - get row vector T(j)
        T_rsum = sum(T,2);  %sum over rows - get column vector T(i)
%Total system throughPUT
    tst_p = sum(sum(T));
P = T'; %flow is column to row for P matrix
        P_rsum=sum(P,2);  %sum over rows - get column vector P(i)
        %P_csum=sum(P,1);  %sum over columns - get row vector P(j)
k = 1;  %select constant multiplier value (usually 1 or 0.7)

%INSTANTANEOUS FRACTIONAL FLOW MATRIX CREATION (Q)
  %INITIALIZING VARIABLES:
    Q=zeros(n_T);    %Instantaneous fractional flow matrix
    i=1;            %counter for the fractional flow matrix creation routine
    %Converts the production matrix to the fractional flow matrix by dividing
    %each element of a nonzero row by the sum of the row's elments
    while i < n_T+1           %loop divides each row of P by its row sum
        if P_rsum(i) > 0            %if statement prevents division when row sum = 0
            Q(i,:) = P(i,:)/P_rsum(i);
        end
        i=i+1;
    end
    
    %TRANSITIVE CLOSURE MATRIX CREATION (N)
    %Calculates the transitive closure matrix using the previously calculated
    %fractional flow matrix => N=(I-Q)^-1
    N = inv(eye(size(P,1))-Q);

    %FLOW METRIC CALCULATION (Warning: this code is problem dependent)
    %Mean path length
    inflow = sum(T(1,:));             %flow entering the system
    internal_flow = sum(sum(T(2:(n_T-2),2:(n_T-2))));
    TST = inflow + internal_flow;     %total system throughflow (inflow + interholon flow)
    MPL = TST/inflow;                  %mean path length
    
    %Mean cyclic path length
    d_N = diag(N);                    %diagonal elements of the N matrix
    c_re = (d_N(2:(n_T-2))-ones(size(d_N(2:(n_T-2))))) ./ d_N(2:(n_T-2));  %return cycling eff. vector
    tst_c = c_re'*P_rsum(2:(n_T-2));
    %pl_c = tst_c/inflow;
    
    %Finn Cycling Index (CI)
   CI = tst_c/TST;

%Average Mutual Information (AMI)
    AMI_ij = zeros(n_T);
    i=1;            %counter
    while i < n_T+1
        j = 1;
            while j < n_T+1   
                AMI_ij(i,j) = (T(i,j)/tst_p)*log2((T(i,j)*tst_p)/(T_rsum(i)*T_csum(j)));
                j = j+1;
            end
        i = i+1;
    end
        AMI_ij(isnan(AMI_ij)) = 0;
        AMI = sum(sum(AMI_ij));

%(ASC) Ascendency 
    ASC = AMI * tst_p; 
      
%Development Capacity (DC)
    DC_ij = zeros(n_T);
    i=1;            %counter
    while i < n_T+1
        j = 1;
            while j < n_T+1   
                DC_ij(i,j) = T(i,j)*log2(T(i,j)/tst_p);
                j = j+1;
            end
        i = i+1;
    end
        DC_ij(isnan(DC_ij)) = 0;
        DC = -1*sum(sum(DC_ij));
       
%Total System Overhead (TSO)
   TSO = DC - ASC;

%Robustness (R)
   R=-1*k*(ASC/DC)*log2(ASC/DC);
   R(isnan(R)) = 0;
   
%Shannon Index (H)
    H_ij = zeros(n_T);
    i=1;            %counter
    while i < n_T+1
        j = 1;
            while j < n_T+1   
                H_ij(i,j) = (T(i,j)/tst_p)*log2(T(i,j)/tst_p);
                j = j+1;
            end
        i = i+1;
    end
        H_ij(isnan(H_ij)) = 0;
        H = -1*k*sum(sum(H_ij));
        
%Average degree of order (alpha)
alpha = (ASC/DC);
end

function [Result]=xlwrite(file,data,sheet)
%% Function xlwrite provides almost the same functionality as the function
% xlswrite, non-accessible from MAC Matlab. It writes numeric or cell data to excel. 
% Syntax is the same as for xlswrite, additional
% Java packages need to be loaded in java path : jxl.jar, MXL.jar and
% matlabcontrol-3.1.0.jar.
%  
% Function handles input and transmits
% data to Java methods in order to write to excel files. Pay attention to sheet name,
% it shouldn't be the same as any other in the excel workbook.
%
% Data is written as text (Label) and a simple click in Excel (Convert to
% number) will solve the issue of numbers seen as a text. 
% For further information, please check Documentation of package MXL Java Classes. 
% 
% Input (3): 
%
%    - filename (char)
%    - data (cell or numerical, up to 3-dimensionnal array) 
%    - sheet name *optional* (char)
%
% Output (1) :
%
%    - Result code, 1 if successful, 0 otherwise
%
% Function calls : WriteXL (JAVA method from package MXL.jar ),
% Cell2JavaString
% 
% See also Cell2JavaString
%
%  Copyright 2012 AAAiC 
%  Date: 2012/04/10

%% INPUT HANDLING

import mymxl.*;
import jxl.*;

if nargin < 3
    
    %% use Contructor with no Sheet
    
 % Convert Type of data : 
    
   
    if isnumeric(data) 
        
        data=num2cell(data);
    elseif iscell(data)
    
        
    
    else
        error('Data to write must be cell or numeric type. Please try again...');
    
    end   
    
    % Size of the matrix data
    
    m=size(data);
    
    if (max(size(m))==2)
        
        WriteXL(java.lang.String(file),Cell2JavaString(data),m(1),m(2));
        Result=1;
    
    elseif (max(size(m))==3)
        
        WriteXL(java.lang.String(file),Cell2JavaString(data),m(1),m(2),m(3));
        Result=1;
    else
        error('data Matrix too large, must have at most 3 dimensions');
    end
    
    
elseif nargin == 3
    
    %% use Constructor with Sheet name
    
        
 % Type of data
    
   
    if isnumeric(data) 
        
        data=num2cell(data);
    elseif iscell(data)
    
    
    else
        error('Data to write must be cell or numeric type. Please try again...');
    end   
    
    % Size of the matrix data
    
    m=size(data);
    
    if (max(size(m))==2)
        WriteXL(java.lang.String(file),Cell2JavaString(data),m(1),m(2),java.lang.String(sheet),exist(file,'file')/2);
        Result=1;
    else
        error('data Matrix too large, when specifying a single sheet, data must have at most 2 dimensions');
    end
    
    
    
else 
    error('Bad number of arguments');
    
    
end

end

function[JavaString]=Cell2JavaString(Glob)
%% Function Cell2JavaString converts  a cell array to a Java String array
% Input (1): 
%
%    - Glob (cell, up to 3-dimensions)
%
% Output (1) :
%
%    - Java String Array of Glob
%
% Function calls : javaArray (JAVA method)
%
% Called by : xlwrite.m
% 
% See also xlwrite
%
%  Copyright 2012 AAAiC 
%  Date: 2012/04/10

sGlob=size(Glob); % size of each dimension

dimGlob=max(size(sGlob)); % number of dimensions of Glob

%% Initialize Java Array of Strings
  if (dimGlob==3) % 3-D Array
    JavaString=javaArray('java.lang.String',sGlob(1),sGlob(2),sGlob(3));
    %% Element-wise define
        for i=1:sGlob(3)
            for j=1:sGlob(2)
                for k=1:sGlob(1)
                    
                    
                    JavaString(k,j,i)=java.lang.String(cell2char(Glob(k,j,i)));
                    
                end   
            end

        end
         

elseif (dimGlob==2) % 2-D Array
    JavaString=javaArray('java.lang.String',sGlob(1),sGlob(2));
    
            for j=1:sGlob(2)
                for k=1:sGlob(1)
                    
                    JavaString(k,j)=java.lang.String(cell2char(Glob(k,j)));
                    
                end
            end

 
   end





end

function S = cell2char(C)
%
% Converts the contents of a cell array of strings into a character 
% array. The contents of the cell C are read element-wise and 
% converted into a char array of length MAXCOL where MAXCOL is 
% the length of the longest string inside the array. 
% Thus the dimensions of the resulting character array S are 
% [NROW, MAXCOL], with NROW being the number of strings in C.
% Strings inside the array that are of shorter length than 
% MAXCOL, are padded with blank spaces, so that S has an homogeneous
% number of columns. In addition, any rows in C that are NaN's
% in IEEE arithmetic representation are replaced by the string 
% 'NaN'.
%
% Syntax: S = CELL2CHAR(C);
%
% Inputs: 
%        c: cell array 
% Outputs:
%        S: character array
%
% See also: mat2cell, num2cell
% 
% Tonatiuh Pena-Centeno
% Created: 08-Jun-10   Last modified: 16-Jun-10
%

% Verifying C has the correct dimensions
if size(C,2) ~= 1
  error('CELL2CHAR: dimensions of input array C dont seem to be correct');
end

% Character array must have a constant number of columns, 
% so they're retrieved first by computing the MAXCOL 
% number from the entire cell C
nRow = size(C,1);
maxCol = 1;
for it = 1:nRow
  Cit = C{it,:};
  % If cell contents is a number, convert to string
  if isnumeric(Cit)
    Cit = num2str(Cit);
  end
  % If cell contents is NaN, then do not take into account
  if isnan(Cit)
    nCol = 0;
  else
    charC = char(Cit);
    nCol = size(charC,2);
  end
  % Updating maxCol
  if nCol > maxCol
    maxCol = nCol;
  end
end  

S = NaN(nRow,maxCol);
for it = 1:nRow
  Cit = C{it,:};
  % If cell contents is numeric, convert to string
  if isnumeric(Cit)
    Cit = num2str(Cit);
  end
  % If cell contents is NaN, then replace with string
  if isnan(Cit)
    Cit = 'NaN';
  end
  charC = char(Cit);
  nCol = size(charC,2);
  eval(['S(it,:) = [charC, blanks(', num2str(maxCol-nCol), ')];']); 
end

% Converting everything into char code
S = char(S);
  
end

function [catmat]=padconcatenation(a,b,c)
%[catmat]=padconcatenation(a,b,c)
%concatenates arrays with different sizes and pads with NaN.
%a and b are two arrays (one or two-dimensional) to be concatenated, c must be 1 for
%vertical concatenation ([a;b]) and 2 for horizontal concatenation ([a b])
%
% a=rand(3,4)
% b=rand(5,2)
% a =
% 
%     0.8423    0.8809    0.7773    0.3531
%     0.2230    0.9365    0.1575    0.3072
%     0.4320    0.4889    0.1650    0.9846
% b =
% 
%     0.6506    0.8854
%     0.8269    0.0527
%     0.4742    0.3516
%     0.4826    0.2625
%     0.6184    0.5161
%
% PADab=padconcatenation(a,b,1)
% PADab =
% 
%     0.8423    0.8809    0.7773    0.3531
%     0.2230    0.9365    0.1575    0.3072
%     0.4320    0.4889    0.1650    0.9846
%     0.6506    0.8854       NaN       NaN
%     0.8269    0.0527       NaN       NaN
%     0.4742    0.3516       NaN       NaN
%     0.4826    0.2625       NaN       NaN
%     0.6184    0.5161       NaN       NaN
%
% PADab=padconcatenation(a,b,2)
% 
% PADab =
% 
%     0.8423    0.8809    0.7773    0.3531    0.6506    0.8854
%     0.2230    0.9365    0.1575    0.3072    0.8269    0.0527
%     0.4320    0.4889    0.1650    0.9846    0.4742    0.3516
%        NaN       NaN       NaN       NaN    0.4826    0.2625
%        NaN       NaN       NaN       NaN    0.6184    0.5161
%
%Copyright Andres Mauricio Gonzalez, 2012.

sa=size(a);
sb=size(b);

switch c
    case 1
        tempmat=NaN(sa(1)+sb(1),max([sa(2) sb(2)]));
        tempmat(1:sa(1),1:sa(2))=a;
        tempmat(sa(1)+1:end,1:sb(2))=b;
        
    case 2
        tempmat=NaN(max([sa(1) sb(1)]),sa(2)+sb(2));
        tempmat(1:sa(1),1:sa(2))=a;
        tempmat(1:sb(1),sa(2)+1:end)=b;
end

catmat=tempmat;
end
