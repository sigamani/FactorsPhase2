import plotly
import plotly.plotly as py
import module1 as m1
import MonthEndChecks as mec

#plotly.tools.set_credentials_file(username='michael.sigamani_sr', api_key='16jtbqurle')

py.sign_in('sigamani1982','lrix3k0xxv')
#py.sign_in('michael.sigamani_sr', 'pUwzlMuStskB3R2zemzc')

#m1.changeName(r"C:\Users\michael\Desktop\SRMA2\\")
#m1.heatmap()

#m1.getMarketsAnalyzerOutput('Identity')

#m1.makePercentiles()
#m1.makeCoverage()
#m1.getCorrelationOfReturns('UNITED KINGDOM')
#m1.makeTrendingStat('Regularity_12Month')

#stat = ['Return_10Year', 'Return_1Year', 'TrackingError','TrackingError_2Year','StdDev','StdDev_2Year','StyleBeta',
#	   'Regularity_3Month','Regularity_6Month','Regularity_12Month','Identity']

#for s in stat:

#    m1.getMarketsAnalyzerOutput(s)



mec.doCountryDataCheck(1.0)

