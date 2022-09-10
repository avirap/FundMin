function ...
fun_PLOT_2(DATES,...
                  PLOT_main,...
                  PLOT_bench,...
                  TITLE,LEGEND,YLABEL)

%% PREAMBLE
    DATES              = [DATES(1)-1;DATES];
    T=length(DATES);

        % EoM
    DATES_temp = fun_last_day_of_month(DATES);
           EoM = find( DATES_temp(:,2)==1 );
           EoY = find( DATES_temp(:,3)==1 );
                   
    PLOT_main_Return   = [ 0 ;100*(PLOT_main(:,2)-1)];
    
        B=size(PLOT_bench,3);
for b=1:B
    PLOT_bench_Return(:,b) = [ 0 ;100*(PLOT_bench(:,2,b)-1)];
    PLOT_bench_StDev(b)    =  100* PLOT_bench(end,4,b)    ;
end


hold on
%% BENCH
for b=1:B
color=[1/3*b,0*b,0*b];
            text(DATES(end)+T/40,PLOT_bench_Return(end,b),['\fontsize{9}',num2str(PLOT_bench_Return(end,b),'%.1f'),'%'],'FontWeight','bold','HorizontalAlignment','left','VerticalAlignment','middle','color',color);
Pbench(1,b)=plot(DATES,PLOT_bench_Return(:,b),    'LineWidth',1.25,'color',color);
Pbench(2,b)=plot(DATES,PLOT_bench_Return(:,b),'o','LineWidth',0.90,'MarkerFaceColor',color,'color','k'); Pbench(2,b).MarkerIndices=EoM;
end

%% MAIN
color=[103/255,142/255,47/255];
           text(DATES(end)+T/40,PLOT_main_Return(end),['\fontsize{9}',num2str(PLOT_main_Return(end),'%.1f'),'%'],'FontWeight','bold','HorizontalAlignment','left','VerticalAlignment','middle','color',color);
Pmain(1,1)=plot(DATES,PLOT_main_Return,    'LineWidth',1.75,'color',color);
Pmain(2,1)=plot(DATES,PLOT_main_Return,'o','LineWidth',1.00,'MarkerFaceColor',color,'color','k');                        Pmain(2).MarkerIndices=EoM;

plot(DATES,zeros(T,1),'k');

%% Y limits
Ymin=min(min([PLOT_main_Return,PLOT_bench_Return]));
Ymax=max(max([PLOT_main_Return,PLOT_bench_Return]));
set(gca,'Ylim',[Ymin,Ymax]);

%% VERTICALS
% fun_event_dates(DATES);
fun_monthly_lines(DATES,EoM);

%% Xticks
Y=length(EoY);
M=length(EoM);

if Y<=2
   for m=2:M
   yr=year(DATES(EoM(m)));
   mnth=month(DATES(EoM(m)));
   dy=day(DATES(EoM(m)));
%    dy=floor(dy/2);
   dtnm=datenum(yr,mnth,dy);
   TICKS(m-1)=find(DATES==dtnm);
   end
   set(gca,'XTick',DATES(TICKS));
   datetick('x', '', 'keepticks','keeplimits');
end
if Y>=3
    for y=2:Y
        yr=year(DATES(EoY(y)));
        mnth=month(DATES(EoY(y)));
        flr=floor(mnth/2);
        dy=eomday(yr,flr);
        dtnm=datenum(yr,flr,dy);
        TICKS(y-1)=find(DATES==dtnm);
    end
    set(gca,'XTick',DATES(TICKS));
    datetick('x','yyyy','keepticks','keeplimits');
end

%% Ytick
Ydiff = Ymax - Ymin;
if 100 < Ydiff && Ydiff < 9999; x=20; set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',12); end
if 50  < Ydiff && Ydiff < 100;  x=10; set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',12); end
if 20  < Ydiff && Ydiff < 50;   x=5;  set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',12); end
if 10  < Ydiff && Ydiff < 20;   x=2;  set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',10); end
if  0  < Ydiff && Ydiff < 10;   x=1;  set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',10); end
ytickformat('%.f') 
%% HORIZONTALS
grid on
ax = gca; ax.GridLineStyle = '-'; ax.GridAlpha = 0.5;

axis tight

%% NO resizing of Font Size
set(gcf, 'PaperPositionMode', 'auto')
%%
title(TITLE,'color',[0/255,2/7,0/255],'FontSize',14,'fontweight','bold');
legend([Pmain(1),Pbench(1,:)],LEGEND,'location','northwest','fontsize',6.5)
% ylabel(YLABEL    ,'FontSize',12,'fontweight','bold');

hold off
end