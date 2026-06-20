#ifndef DATABASECONNECTION_H
#define DATABASECONNECTION_H

#include <string>

class DatabaseConnection {
public:
    DatabaseConnection();
    virtual ~DatabaseConnection();

    void connect();
    void disconnect();
    void executeQuery(std::string sql);

private:
    std::string connectionString;
    bool isConnected;
};

#endif // DATABASECONNECTION_H
