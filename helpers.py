def count_per_NE(df, unit):
    '''
        Return number of NE or cells per hour.
    '''
    df = df.groupby(by=['Start Time'])[unit].nunique().to_frame()
    df.reset_index(inplace=True)
    return df