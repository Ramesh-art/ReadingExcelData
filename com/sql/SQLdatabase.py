import MySQLdb

class MySQLUtils:
    # def __init__(self, user='root', password='Ilovmom@1', host='localhost', database='database1'):
    #     # self.connection = MySQLdb.connect(user=user, password=password, host=host, database=database)
    def __init__(self):
        pass
    def saveData(self, row,user='root', password='Ilovmom@1', host='localhost', database='database1'):
        try:
            connection = MySQLdb.connect(user=user, password=password, host=host, database=database)
            #print(type(row))
            query = """INSERT INTO oms_uc6_search_light(testcase_id,module,                         
                                                    car_id,recording_id,     
                                                    steeringwheel,tripparts,     
                                                    interior,testperformedby,     
                                                    occupiedseats,testcase_description,     
                                                    object_id,evaluation,     
                                                    datfile,    
                                                    falsetriggers,     
                                                    result)
                                                    VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
            recordTup = (row.TestCaseId,
                         row.Module,
                         row.CarId,
                         str(row.RecordingId),
                         row.Steeringwheel,
                         row.TripParts,
                         row.Interior,
                         row.TestPerformedby,
                         row.OccupiedSeats,
                         row.TestcaseDescription,
                         row.ObjectId,
                         row.Evaluation,
                         row.DATfile,
                         row.FalseTriggers,
                         row.Result)
            # UseCaseModel().CarId
            cursor = connection.cursor()
            a = cursor.execute(query,recordTup)
            cursor.close()
            print("Saved Successflly ", a)
        except Exception as error:
            print(error)
        finally:
            cursor.close()
            connection.close()
