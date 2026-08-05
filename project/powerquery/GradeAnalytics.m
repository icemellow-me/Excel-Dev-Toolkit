let
    // ═══════════════════════════════════════════════════════════════
    // Power Query M: Grade Analytics ETL Pipeline
    // ═══════════════════════════════════════════════════════════════
    // Unpivots subject scores for analytics, computes subject-wise
    // and gender-wise statistics.
    // 

    // ── Source: Read from the processed student table ──
    Source = Excel.CurrentWorkbook(){[Name="StudentData"]}[Content],
    
    // ── Select only the columns we need ──
    KeepColumns = Table.SelectColumns(Source, {
        "StudentID", "Full Name", "Gender", "House",
        "English", "Mathematics", "Science", "ICT", "Social Studies"
    }),
    
    // ── Unpivot subject scores into rows ──
    // This transforms: [ST001, Daniel, 85, 78, 82, 90, 79]
    // Into 5 rows: [ST001, Daniel, "English", 85],
    //              [ST001, Daniel, "Mathematics", 78], ...
    Unpivoted = Table.UnpivotOtherColumns(KeepColumns,
        {"StudentID", "Full Name", "Gender", "House"},
        "Subject", "Score"
    ),
    
    // ── Type the new columns ──
    TypedUnpivot = Table.TransformColumnTypes(Unpivoted, {
        {"Subject", type text},
        {"Score", type number}
    }),
    
    // ════ Analysis 1: Subject-Wise Statistics ════
    SubjectStats = Table.Group(TypedUnpivot, {"Subject"},
        {
            {"Average Score", each Number.Round(List.Average([Score]), 2), type number},
            {"Highest", each List.Max([Score]), type number},
            {"Lowest", each List.Min([Score]), type number},
            {"Pass Count", each List.Count(List.Select([Score], each _ >= 50)), Int64.Type},
            {"Fail Count", each List.Count(List.Select([Score], each _ < 50)), Int64.Type},
            {"Student Count", each Table.RowCount(_), Int64.Type}
        }
    ),
    
    // ════ Analysis 2: Gender Comparison ════
    GenderStats = Table.Group(TypedUnpivot, {"Gender"},
        {
            {"Average Score", each Number.Round(List.Average([Score]), 2), type number},
            {"Total Students", each List.Count(List.Distinct([StudentID])), Int64.Type},
            {"Total A Grades", each List.Count(
                List.Select([Score], each _ >= 80)
            ), Int64.Type}
        }
    ),
    
    // ════ Analysis 3: House Comparison ════
    HouseStats = Table.Group(TypedUnpivot, {"House"},
        {
            {"Average Score", each Number.Round(List.Average([Score]), 2), type number},
            {"Student Count", each List.Count(List.Distinct([StudentID])), Int64.Type}
        }
    ),
    
    // ════ Analysis 4: Top 5 Students ════
    StudentTotals = Table.Group(TypedUnpivot, {"StudentID", "Full Name"},
        {
            {"Total Score", each List.Sum([Score]), type number},
            {"Average Score", each Number.Round(List.Average([Score]), 2), type number}
        }
    ),
    Top5 = Table.FirstN(Table.Sort(StudentTotals, {{"Average Score", Order.Descending}}), 5),
    
    // ── Combine all into a single output table ──
    // In practice you'd load each as a separate query
    FinalOutput = SubjectStats
    
in
    FinalOutput
