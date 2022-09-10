function DATES = fun_last_day_of_month(DATES)

DATES_datevec=datevec(DATES);

for t=1:length(DATES_datevec)
    % End of MONTH
    if DATES_datevec(t,3)==eomday(DATES_datevec(t,1),DATES_datevec(t,2))
       DATES(t,2)=1;
    end
    % End of YEAR
    if DATES_datevec(t,3)==31 && DATES_datevec(t,2)==12
       DATES(t,3)=1;
    end
    % LAST day is always end of MONTH & YEAR
    if t==length(DATES)
       DATES(t,2)=1;
       DATES(t,3)=1;
    end
end
