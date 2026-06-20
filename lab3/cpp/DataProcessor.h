#ifndef DATAPROCESSOR_H
#define DATAPROCESSOR_H

#include <string>

class DataProcessor {
public:
    DataProcessor();
    virtual ~DataProcessor();

    void processData();
    void validateInput(std::string data);
};

#endif // DATAPROCESSOR_H
