%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Maldi processing %%%%%%%%%%%
%% By: Mohammed H. Sarikahya %%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%% Purpose: Generate all the neccessary calculations for Area under the curve calculation
%%          including Peak assignment and calculation.
%%          Comparisons of the internally generated peak list and one by the software (e.g., MSIreader or others)
%%          will reveal identical values, if not less Peak assignments. MSIreader (and others) do not auto
%%          generate AUC's, and instead provides intensity values, which are highly experimentally
%%          variable. AUC's provide a more robust assessment of the 'amount' or 'abundance' of a given metabolite.

%% Dependencies: MATLAB ver 2019a and above 
%%               Signal Processing Toolbox (Addon within Matlab)

%% File Type: Must include the Average Spectrum for all available m/z within a sample selection;
%%            Software used to access MALDI IMS images does NOT matter, as long as the spat out
%%            values are organized with "m/z" in col A, and "intenisty" in col B so:
%%                                                  m/z             intensity
%%                                                80.124124        12.4124124
%%                                                80.125332        10.6232232
%%                             Must have all the values present for the selection (e.g., typically > 10k m/z)      
%%                          
 
%% USAGE:
% Before using this script, create a directory for 'Excels' and a directory
% for 'Mats', with those exact cases

% Next, put all Maldi excels in one folder, or seperate by brain region, it
% doesnt matter, since you will need to set the brain region within the
% next few steps

% Read the lines 31 to 57, and fill in the places with 'EDIT'.
% Then hit the F5 button on your keyboard to Run the script, or click Run in
% the Editor tab above

function [] = MALDI_processing()          %do not touch
clc;                                      %do not touch
clear;                                    %do not touch

% Set these Directories for the Script to work
DataDirectory = 'E:\2021_THC\Adult_Quantitation\'; %EDIT'; contains the MALDI excel sheets, and the excel and mats directories
ExcelDirectory = 'E:\2021_THC\Adult_Quantitation\Excels\'; %EDIT'; Excel dir is created within script, but needs a location you want it at
MatsDirectory = 'E:\2021_THC\Adult_Quantitation\Mats\'; %EDIT'; MAT dir is created within script, but needs a location you want it at
                  % Script does work on MACs; just switch direction of "\" to "/"

% Set the name of the second sheet in the excel file extracted from MSireader
SecondSheetName = 'Average Spectrum'; %EDIT

% Set the name of the sheet where you will save the data to:
Datasheet = 'Sheet1'; 

% Set the brain region of interest. Can signify Right/Left as well, if
% needed, e.g., ('*Right DS*') or ('*Left DS*'), but just, e.g., ('*DS*'); 
% or just use ('*xlsx*') and do the calculation for everything in folder selected
  
cd(fullfile(DataDirectory)); %do not touch

BrainRegion = dir('*xlsx*');  % EDIT' Make sure to maintain the asterisk *, otherwise it will not work
                                                                           
  
%% %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%% DO NOT ALTER ANYTHING BELOW THIS LINE %%%%%%%%%%%%%%%%%%%%%%%%%%%%
%-------------------------------------------------------------------------%
%% Start of Script %

mkdir Excels
mkdir Mats
  
DATADIR = DataDirectory; 
ExcelDIR = ExcelDirectory; 
MatsDIR = MatsDirectory; 

ps = dir(strcat(DATADIR));
ps = {ps.name};
ps = ps(~ismember(ps, {'.', '..', '.DS_Store'}));


for i = 1:size(ps,2)
    cd(fullfile(DATADIR));
    tmpfile = BrainRegion; 
    tmpfile = {tmpfile.name};
    tmpfile = tmpfile(~ismember(tmpfile,{'.','..','.DS_Store'}));
end

for i = 1:size(tmpfile,2)
    
    file_temp = char(tmpfile(1,i));  
    match = '.xlsx';
    file = erase(file_temp,match);
    filename = strcat(DATADIR, file, '.xlsx');
    [numeric, text, raw] = xlsread(filename, SecondSheetName); 
    
%     numeric(1:3000,:) = []; % Remove the the first rows till ~80 m/z
%     numeric(:,3:10) = []; % Remove the empty coloumns

    % set coloumn 1 and 2 to X and Y; use numeric, to filter out names/other

    X =  numeric(:,1);
    Y =  (numeric(:,2)); 
% 
%     Y =  normc(numeric(:,2)); %Deep learning Toolbox, norm attempt


    [pks,locs,w,p] = findpeaks(Y);
    %% Plots to check data
    % findpeaks(Y); %Use findpeaks without output arguments to display the peaks
    % plot(X,Y) %plots m/z as x value, and intensity as y value
    %% Find min and max of the range of each peak
    for k = 1:length(locs)
       i = locs(k);
       while i > 1 && Y(i-1) <= Y(i)
           i = i - 1;
       end
       b_min(k) = i;
       i = locs(k);
       while i < length(Y) && Y(i+1) <= Y(i)
           i = i + 1;
       end
       b_max(k) = i;
    end

    %k: number of peaks
    %i: number of points
    %b_min: minimum boundary
    %b_max: maximum boundary


    %% Change in m/z

    for x = 1
        count = 1;
        temp = 0; 
        X =  vertcat(X, temp); 
        for k = 1:length(locs)
               A = (X([b_min(k):b_max(k)+1])); 
               for p = 1:length(A)-1 
                    if p < length(A)
                        A_x = [A(p,:)];
                        A_p = [A(p+1,:)];
                    end
                 change(p) = (A_p - A_x);
               end
        change(end)=[];
        change_in_mz{k}(:,:) = (change(1,:));
        change = 0;
        end
    end
    change(end)=[];

    %Peak m/z values
    for x = 1
        for k = 1:length(locs)
            i = locs(k);
               for p = i
                   t = X(p);
               end
             peakmz(k) = t;
        end

    end

    %% Change in intensity
    for y = 1
        count = 1;
        temp = 0; 
        Y =  vertcat(Y, temp); 
        for k = 1:length(locs)
               A = (Y([b_min(k):b_max(k)+1])); 
               for p = 1:length(A)-1 
                    if p < length(A)
                        A_y = [A(p,:)];
                        A_p = [A(p+1,:)]; 
                    end
                 change(p) = (A_p + A_y)/2;
               end
        change(end)=[];
        change_in_intensity{k}(:,:) = (change(1,:));
        change = 0;
        end
    end

    change_in_intensity = change_in_intensity'; %'
    change_in_mz = change_in_mz'; %'

    %% Intensity subarea (rectangles)

    for m = 1
        count = 1;
        temp = 0;
        change_in_intensity =  vertcat(change_in_intensity, temp); 
        change_in_mz = vertcat(change_in_mz, temp);
        for k = 1
            for j = 1:length(pks)
                for d = 1:length(change_in_intensity{j,1}+1)
                    if d < length(change_in_intensity{j,1})
                        ci = (change_in_intensity{j,1}(d));
                        cm = (change_in_mz{j,1}(d));
                    elseif d == length(change_in_intensity{j,1}-1)
                        ci = (change_in_intensity{j,1}(length(change_in_intensity{j,1})));
                        cm = (change_in_mz{j,1}(length(change_in_mz{j,1})));
                    end
                subarea(d) = cm*ci;
                end
            avg_intensity_subarea{1,j} = (subarea(1,:));
            subarea = 0;
            end
        end
    end
    
    change_in_mz = [];
    change_in_intensity = [];
    
    %% Total Peak Widths (m/z)

    for k = 1:length(locs)
        if (b_min(k) : b_max(k))
            A = (X([b_min(k)]));
            B = (X([b_max(k)]));
            i = B - A;
        elseif k <= 1:length(locs)
            break;
        end
        total_peak_width(k) = i;
    end

    %% Average Background Intensity

    for j = 1:length(locs)
        if (b_min(j) : b_max(j))
            A = (Y([b_min(j)]));
            B = (Y([b_max(j)]));
            u = (B + A)/2;
        elseif j <= 1:length(locs)
            break;
        end
        inten_Y(j) = u;
    end


    %% Average Background to Extract from sum of peak area under curve
    for k = 1:length(locs)
        avg_back(k) = inten_Y(k) * total_peak_width(k);
    end

    %% sum of subarea intensity
    avg_intensity_subarea = avg_intensity_subarea'; %'

    for k = 1:length(avg_intensity_subarea)
        sum_AUC(k) = sum(avg_intensity_subarea{k,1});
    end

    %% SAMPLE PEAK AUC %%

    for k = 1:length(locs)
        SAMPLE_PEAK_AUC(k) = sum_AUC(k) - avg_back(k);
    end


    %%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    %%% Save the values %%%%%%%%%%%%%%%%%%%%%%%%%
    %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

    %%%%%%%%  Save in .mat  %%%%%%%%%%%%
    peakmz = peakmz'; %'
    total_peak_width = total_peak_width'; %'
    inten_Y = inten_Y'; %'
    avg_back = avg_back'; %'
    sum_AUC = sum_AUC'; %'
    SAMPLE_PEAK_AUC = SAMPLE_PEAK_AUC'; %'

    save(strcat(DATADIR, file, '_ANALYZED'),'peakmz');
    save(strcat(DATADIR, file, '_ANALYZED'),'pks', '-append');
    save(strcat(DATADIR, file, '_ANALYZED'),'total_peak_width', '-append');
    save(strcat(DATADIR, file, '_ANALYZED'),'inten_Y', '-append');
    save(strcat(DATADIR, file, '_ANALYZED'),'avg_back', '-append');
    save(strcat(DATADIR, file, '_ANALYZED'),'sum_AUC', '-append');
    save(strcat(DATADIR, file, '_ANALYZED'),'SAMPLE_PEAK_AUC', '-append');
    load(strcat(DATADIR, file, '_ANALYZED'));

    %%%%%%%%  Save in .xlsx  %%%%%%%%%%%%

    col_header={'Peak m/z','Peak Instensity','Total Peak Widths (m/z)','Average Background Intensity','Average Background to Extract from sum of peak area under curve','sum of peak area under curve (AUC)', 'SAMPLE PEAK AUC'};     %Row cell array (for column labels)
    xlswrite(strcat(DATADIR, file, '_ANALYZED'), col_header,Datasheet,'A1');     %Write column header

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), peakmz, Datasheet, 'A2');

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), pks, Datasheet, 'B2');

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), total_peak_width, Datasheet, 'C2');

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), inten_Y, Datasheet, 'D2');

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), avg_back, Datasheet, 'E2');

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), sum_AUC, Datasheet, 'F2');

    xlswrite(strcat(DATADIR, file, '_ANALYZED'), SAMPLE_PEAK_AUC, Datasheet, 'G2');
    
      % Move the files to the correct directory
      movefile((strcat(DATADIR, file, '_ANALYZED', '.xls')),ExcelDIR)
      movefile((strcat(DATADIR, file, '_ANALYZED', '.mat')),MatsDIR)
      
      % Reset the variables to be filled for loop
        peakmz = []; %'
        total_peak_width = []; %'
        inten_Y = []; %'
        avg_back = []; %'
        sum_AUC = []; %'
        SAMPLE_PEAK_AUC = []; %'
        avg_intensity_subarea = [];
end
end
