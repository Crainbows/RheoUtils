
% Open dialog to get file
[file,path] = uigetfile({'*.xlsx;*.xls'});

% Did the user select a file?
if isequal(file,0)
   disp('No file selected.');
else
    
%     Get the sheets from the excel file
    [status,sheets] = xlsfinfo(fullfile(path,file));
    
%     Which sheets does the user want to process
    [indx,tf] = listdlg('PromptString','Select Sheets:',...
                           'ListString',sheets);

%     Linear Plot G' and G'' vs Temperature
    fig = figure('Name',file,'NumberTitle','off');
    fig.Color = 'w';
    
    linesty = {'-','--','-.',':'};
    
    hold on;

%     Cycle through sheets and plot on same graph
    for k=1:length(indx)
        [num,txt,raw] = xlsread(fullfile(path,file),indx(k));

        plot(num(:,7),smooth(num(:,1)),...
            'Color', 'r',...
            'LineStyle', linesty{k},...
            'LineWidth',2); % Storage Mod
        
        plot(num(:,7),smooth(num(:,2)),...
            'Color', 'b',...
            'LineStyle', linesty{k},...
            'LineWidth',2); % Loss Mod
    end

%     Setting legend properties
    legend('G'' Heating','G'''' Heating','G'' Cooling','G'''' Cooling');
    legend('boxoff');

%     Set Axis properties
    ax = gca;
    ax.LineWidth = 1.5;
    ax.FontSize = 16;

%     Set Graph Labels
    xlabel('Temperature (\circC)');
    ylabel('G'',G''''(Pa)');
    
%     Create Dropdown UI element
    popup = uicontrol('Style', 'popup',...
       'String', {'Linear','Log'},...
       'Position', [450 360 100 50],...
       'Callback', @setscale);
   
%     Create push button UI element  
    btn = uicontrol('Style', 'pushbutton', 'String', 'Print',...
        'Position', [390 390 50 20],...
        'Callback', @printimg);   

end

% Clear unused vars
clearvars sheets tf txt num raw status k

function setscale(source,~) % Event not used
    val = source.Value;
    maps = source.String;
    ax = gca;
    ax.YScale = maps{val};
    switch val
        case 1
            yticks('auto');
        case 2
            yticks([0.0001 0.001 0.01 0.1 1 10 100 1000 10000 100000 1000000  10000000 100000000]);
    end
    disp(strcat("Scale set to ", maps{val}));
end

function printimg(~,~) % Source and Event not used
    [file,path,indx] = uiputfile({'*.fig';'*.png';'*.emf';'*.svg'});
    if isequal(file,0)
        disp('Cancelled Save.');
    else
        fig = gcf;
        btn = findobj('Style','push');
        btn.Visible = 'off';
        popup = findobj('Style','popup');
        popup.Visible = 'off';
        disp("Saving...");
        switch indx
            case 1
                saveas(fig,fullfile(path,file)); % Matlab format
            case 2
                print(fullfile(path,file), '-dpng','-r600');
            case 3
                print(fullfile(path,file), '-demf','-r600');
            case 4
                print(fullfile(path,file), '-dsvg','-r600');
        end
        disp("Finished.");
        popup.Visible = 'on';
        btn.Visible = 'on';
    end
end




