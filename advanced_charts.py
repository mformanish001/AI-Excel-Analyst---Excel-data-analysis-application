"""Interactive chart builder for AI Excel Analyst."""
import io
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from uuid import uuid4

TYPES = ['Bar','Horizontal Bar','Stacked Bar','100% Stacked Bar','Line','Step Line','Area','Stacked Area','Pie','Donut','Scatter','Bubble','Histogram','Box','Violin','Density','Heatmap','Correlation Heatmap','Waterfall','Funnel','Treemap','Sunburst','Sankey','Radar','Polar','Lollipop','Dumbbell','Gauge','Bullet','KPI','Candlestick','OHLC','Calendar Heatmap']
AGGS = {'Sum':'sum','Average':'mean','Count':'count','Minimum':'min','Maximum':'max','None':None}

def numeric(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors='coerce')
    return pd.to_numeric(s.astype('string').str.replace(r'[$₹€,]', '', regex=True).str.strip(), errors='coerce')

def build_chart(df, kind, x, y, agg='Sum', group=None, size=None, top_n=50, secondary=None, source=None, target=None, date=None, open_col=None, high=None, low=None, close=None, goal=None):
    if df is None or df.empty: raise ValueError('No rows available.')
    if x not in df or y not in df: raise ValueError('Selected X/Y columns are not present. Please select available columns.')
    data = df.copy()
    value_needed = kind not in {'KPI','Pie','Donut','Funnel','Treemap','Sunburst','Sankey','Calendar Heatmap'} or y is not None
    if y and kind not in {'Candlestick','OHLC'}:
        data[y] = numeric(data[y])
        if data[y].notna().sum()==0: raise ValueError(f'"{y}" has no usable numeric values. Convert its data type first or select a numeric column.')
    if kind in {'Candlestick','OHLC'}:
        cols=[date,open_col,high,low,close]
        if not all(c in df.columns for c in cols): raise ValueError('Candlestick/OHLC needs Date, Open, High, Low and Close columns.')
        for c in [open_col,high,low,close]: data[c]=numeric(data[c])
        data=data.dropna(subset=cols).sort_values(date)
        trace = go.Candlestick if kind=='Candlestick' else go.Ohlc
        fig=go.Figure(data=[trace(x=data[date],open=data[open_col],high=data[high],low=data[low],close=data[close])])
        return fig,data
    if kind=='Correlation Heatmap':
        nums=df.apply(numeric).dropna(axis=1,how='all')
        if nums.shape[1]<2: raise ValueError('At least two numeric columns are needed for correlation.')
        corr=nums.corr()
        return px.imshow(corr,text_auto='.2f',aspect='auto',title='Correlation heatmap'),corr.reset_index()
    if kind=='Sankey':
        if not source or not target: raise ValueError('Select source and target columns.')
        temp=data[[source,target,y]].dropna().groupby([source,target],as_index=False)[y].sum()
        labels=list(dict.fromkeys(temp[source].astype(str).tolist()+temp[target].astype(str).tolist()))
        ids={v:i for i,v in enumerate(labels)}
        fig=go.Figure(go.Sankey(node=dict(label=labels),link=dict(source=temp[source].astype(str).map(ids),target=temp[target].astype(str).map(ids),value=temp[y])))
        return fig,temp
    if kind=='Calendar Heatmap':
        d=date or x
        dt=pd.to_datetime(data[d],errors='coerce')
        temp=pd.DataFrame({'date':dt,'value':data[y]}).dropna()
        if temp.empty: raise ValueError('No valid dates found for calendar heatmap.')
        temp['week']=temp.date.dt.isocalendar().week.astype(int)
        temp['weekday']=temp.date.dt.day_name()
        temp=temp.groupby(['weekday','week'],as_index=False)['value'].sum()
        return px.density_heatmap(temp,x='week',y='weekday',z='value',histfunc='sum',title='Calendar heatmap'),temp
    if kind in {'Histogram','Box','Violin','Density'}:
        temp=data.dropna(subset=[y])
        if kind=='Histogram': fig=px.histogram(temp,x=y,color=group)
        elif kind=='Box': fig=px.box(temp,x=x if x!=y else None,y=y,color=group)
        elif kind=='Violin': fig=px.violin(temp,x=x if x!=y else None,y=y,color=group,box=True)
        else: fig=px.histogram(temp,x=y,histnorm='probability density',marginal='rug',color=group)
        return fig,temp[[c for c in [x,y,group] if c is not None]].head(5000)
    if kind in {'Scatter','Bubble'}:
        data[x]=numeric(data[x])
        if data[x].notna().sum()==0: raise ValueError('Scatter and bubble charts need a numeric X-axis.')
        if kind=='Bubble':
            if not size: raise ValueError('Select a bubble size column.')
            data[size]=numeric(data[size]).clip(lower=0)
        temp=data.dropna(subset=[x,y])
        fig=px.scatter(temp,x=x,y=y,color=group,size=size if kind=='Bubble' else None,hover_data=list(df.columns[:4]),trendline=None)
        return fig,temp.head(5000)
    if kind in {'Gauge','Bullet','KPI'}:
        value=float(data[y].dropna().sum() if agg=='Sum' else data[y].dropna().mean())
        if kind=='KPI': fig=go.Figure(go.Indicator(mode='number',value=value,title={'text':y}))
        elif kind=='Gauge': fig=go.Figure(go.Indicator(mode='gauge+number',value=value,title={'text':y},gauge={'axis':{'range':[0,max(value*1.2,goal or 1)]},'threshold':{'value':goal} if goal else {}}))
        else: fig=go.Figure(go.Indicator(mode='number+gauge',value=value,title={'text':y},gauge={'shape':'bullet','axis':{'range':[0,max(value*1.2,goal or 1)]},'threshold':{'value':goal} if goal else {}}))
        return fig,pd.DataFrame({y:[value]})
    keys=[x]+([group] if group and group!=x else [])
    if agg=='None':
        temp=data[keys+[y]].dropna().head(top_n)
    else:
        temp=data.dropna(subset=[y]).groupby(keys,dropna=False,as_index=False)[y].agg(AGGS[agg])
        temp=temp.sort_values(y,ascending=False).head(top_n)
    if temp.empty: raise ValueError('No valid rows available for this chart.')
    if kind in {'Bar','Stacked Bar','100% Stacked Bar'}:
        fig=px.bar(temp,x=x,y=y,color=group,barmode='stack' if kind!='Bar' else 'group')
        if kind=='100% Stacked Bar': fig.update_layout(barnorm='percent')
    elif kind=='Horizontal Bar': fig=px.bar(temp,x=y,y=x,color=group,orientation='h')
    elif kind in {'Line','Step Line'}:
        temp=temp.sort_values(x)
        fig=px.line(temp,x=x,y=y,color=group,markers=True)
        if kind=='Step Line': fig.update_traces(line_shape='hv')
    elif kind in {'Area','Stacked Area'}: fig=px.area(temp.sort_values(x),x=x,y=y,color=group,groupnorm=None)
    elif kind in {'Pie','Donut'}: fig=px.pie(temp,names=x,values=y,hole=0.5 if kind=='Donut' else 0)
    elif kind=='Heatmap':
        if not group: raise ValueError('Choose a Group/Series column for the heatmap.')
        pivot=temp.pivot_table(index=x,columns=group,values=y,aggfunc='sum')
        fig=px.imshow(pivot,aspect='auto',title=f'{y} by {x} and {group}')
    elif kind=='Waterfall': fig=go.Figure(go.Waterfall(x=temp[x].astype(str),y=temp[y],measure=['relative']*len(temp)))
    elif kind=='Funnel': fig=px.funnel(temp,x=y,y=x)
    elif kind=='Treemap': fig=px.treemap(temp,path=[x]+([group] if group else []),values=y)
    elif kind=='Sunburst': fig=px.sunburst(temp,path=[x]+([group] if group else []),values=y)
    elif kind in {'Radar','Polar'}:
        fig=px.line_polar(temp,r=y,theta=x,color=group,line_close=True) if kind=='Radar' else px.bar_polar(temp,r=y,theta=x,color=group)
    elif kind=='Lollipop':
        fig=go.Figure()
        for _,r in temp.iterrows(): fig.add_shape(type='line',x0=str(r[x]),x1=str(r[x]),y0=0,y1=float(r[y]),line=dict(color='lightgray'))
        fig.add_trace(go.Scatter(x=temp[x].astype(str),y=temp[y],mode='markers',marker_size=11))
    elif kind=='Dumbbell':
        if not secondary or secondary not in data: raise ValueError('Choose a second numeric value column for Dumbbell.')
        other=data.groupby(x,as_index=False)[secondary].agg(AGGS[agg] or 'mean').head(top_n)
        temp=temp.merge(other,on=x,how='left')
        fig=go.Figure()
        for _,r in temp.iterrows(): fig.add_trace(go.Scatter(x=[r[y],r[secondary]],y=[str(r[x])]*2,mode='lines+markers',showlegend=False))
        fig.update_layout(xaxis_title=f'{y} / {secondary}',yaxis_title=x)
    else: raise ValueError('Unsupported chart type: '+kind)
    fig.update_layout(title=f'{kind}: {y} by {x}',template='plotly_white',margin=dict(t=65,b=65))
    return fig,temp

def render_advanced_builder(df):
    st.subheader('📈 Advanced Chart Builder')
    st.caption('Manual column selection is authoritative. Charts are interactive; unavailable fields produce an error rather than being silently substituted.')
    if df is None or df.empty: return
    cols=list(df.columns)
    def initial(name, fallback=0): return cols.index(name) if name in cols else fallback
    kind=st.selectbox('Chart type',TYPES,key='adv_chart_kind')
    c1,c2,c3=st.columns(3)
    with c1: x=st.selectbox('X-axis / category',cols,index=initial('movie_id',initial('movie_id_')),key='adv_chart_x')
    with c2: y=st.selectbox('Y-axis / numeric value',cols,index=initial('Budget',initial('budget')),key='adv_chart_y')
    with c3: agg=st.selectbox('Aggregation',list(AGGS),key='adv_chart_agg')
    group=None;size=None;secondary=None;source=None;target=None;date=None;open_col=None;high=None;low=None;close=None;goal=None
    with st.expander('Advanced chart options',expanded=kind in {'Sankey','Candlestick','OHLC','Bubble','Dumbbell','Heatmap'}):
        group=st.selectbox('Group / series / color',['(None)']+cols,key='adv_chart_group'); group=None if group=='(None)' else group
        top_n=st.slider('Maximum categories / rows',5,500,50,key='adv_chart_top')
        if kind=='Bubble': size=st.selectbox('Bubble size',cols,key='adv_chart_size')
        if kind=='Dumbbell': secondary=st.selectbox('Second value',cols,key='adv_chart_second')
        if kind=='Sankey':
            source=st.selectbox('Source',cols,key='adv_chart_source')
            target=st.selectbox('Target',cols,index=min(1,len(cols)-1),key='adv_chart_target')
        if kind in {'Candlestick','OHLC'}:
            date=st.selectbox('Date',cols,key='adv_chart_date')
            open_col=st.selectbox('Open',cols,key='adv_chart_open')
            high=st.selectbox('High',cols,key='adv_chart_high')
            low=st.selectbox('Low',cols,key='adv_chart_low')
            close=st.selectbox('Close',cols,key='adv_chart_close')
        if kind=='Calendar Heatmap': date=st.selectbox('Date column',cols,key='adv_chart_calendar_date')
        if kind in {'Gauge','Bullet'}: goal=st.number_input('Target (optional, 0 for automatic)',value=0.0,key='adv_chart_goal') or None
    try:
        fig,chart_data=build_chart(df,kind,x,y,agg,group,size,top_n,secondary,source,target,date,open_col,high,low,close,goal)
    except Exception as exc:
        st.warning(str(exc)); return
    st.plotly_chart(fig,use_container_width=True)
    st.caption(f'Using {x} and {y} • {agg} • {len(chart_data):,} plotted records/categories')
    html=fig.to_html(full_html=True,include_plotlyjs=True).encode('utf-8')
    st.download_button('Download interactive HTML',html,file_name='advanced_chart.html',mime='text/html')
    with st.expander('View chart data'): st.dataframe(chart_data,use_container_width=True)
    if st.button('Save Visualization to Report',type='primary',key='adv_chart_save'):
        item={'id':uuid4().hex[:10],'type':'chart','chart_type':kind,'name':f'{kind}: {y} by {x}','request':f'Manual {kind}; {agg} {y} by {x}','plotly_json':fig.to_json(),'plotly_html':html,'data':chart_data.copy()}
        st.session_state.visualizations.append(item)
        st.session_state.report_items=[v.copy() for v in st.session_state.visualizations]
        st.success('Interactive chart saved to Visualizations. Download HTML to retain interactivity.')
