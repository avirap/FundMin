% Y=fun_dates(X,Freq,in_1,in_2,out_1,out_2)
%
% X      is Input Date Vector
% Freq   is D,M or Q. It is Q ONLY if input is format yyyyqq i.e 1947q4.
%
% in_1 is datenum/double/string/cell
% in_2 is format 'yyyymmdd' etc...
%
% out_1 is datenum/double/string
% out_2 is format 'yyyymmdd' etc...
%
% OUTPUT FORMAT FOR 'M' HAS ALWAYS DD=01
% OUTPUT FORMAT FOR 'Q' HAS ALWAYS MM=03 || 06 || 09 ||12 and DD=01
% Output format for 'D' is flexible.


function Y=fun_dates(X,Freq,in_1,in_2,out_1,out_2)
if strcmp(Freq,'M')
   if     strcmp(in_1,'datenum')
          
          X=datestr(X,'yyyy-mm');
          X=datenum(X,'yyyy-mm');
          
          if     strcmp(out_1,'datenum')
          
                 Y=X;
                 
          elseif strcmp(out_1,'double')
                 
                 X=datestr(X,out_2);
                 Y=str2num(X);
                          
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end      
   elseif strcmp(in_1,'double')
          
          X=num2str(X);
          X=datenum(X,in_2);
          
          X=datestr(X,'yyyy-mm');
          X=datenum(X,'yyyy-mm');
          
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end
   elseif strcmp(in_1,'string')
          
          X=datenum(X,in_2);
          
          X=datestr(X,'yyyy-mm');
          X=datenum(X,'yyyy-mm');
          
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end
   elseif strcmp(in_1,'cell')
       
          X=char(X);
          X=datenum(X,in_2);
          
          
          X=datestr(X,'yyyy-mm');
          X=datenum(X,'yyyy-mm');
          
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end    
   end
elseif strcmp(Freq,'Q')
       if strcmp(in_1,'string')
          
          date_vec=zeros(length(X),6);
          date_vec(:,1)=str2num(X(:,1:4));
          date_vec(:,2)=str2num(X(:,6));
          date_vec(:,2)=3*(date_vec(:,2)-1)+1;
          date_vec(:,3)=ones(length(X),1);
           
          X=datenum(date_vec);
                   
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end
   elseif strcmp(in_1,'cell')
       
          X=char(X);
          date_vec=zeros(length(X),6);
          date_vec(:,1)=str2num(X(:,1:4));
          date_vec(:,2)=str2num(X(:,6));
          date_vec(:,2)=3*date_vec(:,2);
          date_vec(:,3)=ones(length(X),1);
           
          X=datenum(date_vec);
                   
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end    
       end
elseif strcmp(Freq,'D')
      if     strcmp(in_1,'datenum')
          
          if     strcmp(out_1,'datenum')
          
                 Y=X;
                 
          elseif strcmp(out_1,'double')
                 
                 X=datestr(X,out_2);
                 Y=str2num(X);
                          
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end      
     elseif strcmp(in_1,'double')
          
            X=num2str(X);
          X=datenum(X,in_2);
                  
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end
   elseif strcmp(in_1,'string')
          
          X=datenum(X,in_2);
          
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end
   elseif strcmp(in_1,'cell')
       
          X=char(X);
          X=datenum(X,in_2);          
                
          if     strcmp(out_1,'datenum')
                 
                 Y=X;
                 
          elseif strcmp(out_1,'double')
              
                 X=datestr(X,out_2);
                 Y=str2num(X);
                 
          elseif strcmp(out_1,'string')
                 
                 Y=datestr(X,out_2);
          end    
   end
    
end
end

