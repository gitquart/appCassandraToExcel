import database as db


query='select id from thesis.impi_docs_masters where year=2020 ALLOW FILTERING'
db.getLargeQuery(query)



