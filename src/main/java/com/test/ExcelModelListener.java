package com.test;

import com.alibaba.excel.metadata.BaseRowModel;
import com.test.ExcelModel;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

import java.util.ArrayList;
import java.util.List;

public class ExcelModelListener<T extends BaseRowModel> extends AnalysisEventListener<T> {
    private final List<T> rows = new ArrayList<>();

    @Override
    public void invoke(T object, AnalysisContext context) {
        rows.add(object);
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        System.out.println("read {} rows %n" + rows.size());
    }

    public List<T> getRows() {
        return rows;
    }
}
