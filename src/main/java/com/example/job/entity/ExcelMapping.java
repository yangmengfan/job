package com.example.job.entity;

import lombok.Data;
import lombok.experimental.Accessors;

@Data
@Accessors(chain =true)
public class ExcelMapping {
    private String source;
    private int sourceCol;
    private String destination;
    private int destinationCol;

}
