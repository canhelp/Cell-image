clc;clear all;close all;

folder_name = "C:\Users\zhang\OneDrive - Wentworth Institute of Technology\Ming\Software code\Cell image\Image samples";%<----update folder name, give the full path, e.g., c:\image\cell image\

calibration = 10;%<---- update the parameter here to get green cell more counts or less counts (e.g., 10 -- more count; 20--less count)
%You may use the images you already counted before to decide this
%'calibration' parameter. If program generated # generally smaller than
%your counts, change 'calibration' to a smaller number (e.g., 5 or 3). If the
%program generated # bigger than your counts, change 'calibration to a
%bigger number (e.g., 15 or 20)


excel_file = 'cell_count.xls';% <-------update the output excel file name if you want
Thd_red = 160;  %<--------- update red channel threshold



%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
file_name = dir(folder_name);
%has_subfolder = 0; %flag to check subfolder

%create excel file
exl_title={'file name','red count','green count'};
if(exist(strcat(folder_name,'\',excel_file),'file'))
    delete(strcat(folder_name,'\',excel_file));
end

if(~xlswrite(strcat(folder_name,'\','cell_count.xls'),exl_title,'main','A1'))
    disp('Cannot create cell_count.xls file');
end
num = 0;
for file_list = 3:length(file_name)
    
    %     if(isfolder(file_name(file_list).name)   %it is a folder not file
    %         sub_folder_name = strcat(folder_name,'\','file_name(file_list).name')  %move to that folder
    %         sub_file_name = dir(sub_folder_name);
    %     end
    
    if(isempty(strfind(file_name(file_list).name,'.tif'))) %not .tif file, ignore it
        continue;
    end
    
    I = imread(strcat(folder_name,'\',file_name(file_list).name));
    num = num+1;
    
    [red_count,green_count,I_edge] = countcell(I,Thd_red,calibration);
    
    
    %write result to excel
    exl_row ={file_name(file_list).name,red_count,green_count};
    xlswrite(strcat(folder_name,'\',excel_file),exl_row,'main',strcat('A',int2str(num+1)));
    
    %save edge image
    imwrite(I_edge,strcat(folder_name,'\',file_name(file_list).name(1:length(file_name(file_list).name)-4),'.png'),'png');
    
    fprintf('%s     green: %d; red: %d \n',file_name(file_list).name, green_count, red_count);
    
end





function [red_count,green_count,I] = countcell(I,Thd_red,calibration)
[sizeX,sizeY,sizeZ] = size(I);
Ired = zeros(sizeX,sizeY);
Igreen  = I(:,:,2);

%set red and green channel image
for i = 1:sizeX
    for j = 1:sizeY
        if(I(i,j,1)>Thd_red)
            Ired (i,j) = 1;
            Igreen(i,j) = 0;
            
        end
        
    end
end

%count red%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Ired = bwareaopen(Ired,5);
cc_red = bwconncomp(Ired,8);
red_count = cc_red.NumObjects;

Ired_edge = edge(Ired,'sobel');%get red cell boundary

%draw red cell boundary as red
for i = 1:sizeX
    for j = 1:sizeY
        if(Ired_edge(i,j)>0)
            I(i,j,1) = 255;
            I(i,j,2) = 0;
            I(i,j,3) = 0;
        end
    end
end


%count green%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Igreen = imextendedmax(Igreen,calibration,8);  
Igreen = bwareaopen(Igreen,7);  % mask size to remove small noise 
cc_green = bwconncomp(Igreen,8);
green_count = cc_green.NumObjects;

Igreen_edge = edge(Igreen,'sobel');%get green cell boundary

%draw green cell boundary as white
for i = 1:sizeX
    for j = 1:sizeY
        if(Igreen_edge(i,j)>0)
            I(i,j,1) = 255;
            I(i,j,2) = 255;
            I(i,j,3) = 255;
        end
    end
end


end