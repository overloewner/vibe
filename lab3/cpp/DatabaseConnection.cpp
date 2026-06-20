#include "DatabaseConnection.h"

DatabaseConnection::DatabaseConnection() : isConnected(false) {}
DatabaseConnection::~DatabaseConnection() {}

void DatabaseConnection::connect() {}
void DatabaseConnection::disconnect() {}
void DatabaseConnection::executeQuery(std::string sql) {}
