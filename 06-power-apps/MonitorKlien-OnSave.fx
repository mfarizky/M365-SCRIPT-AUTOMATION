If(
    // ==========================================================================
    // 0. CEK RACE CONDITION: apakah data sudah diubah orang lain sejak dibuka?
    // ==========================================================================
    !IsBlank(varLastModifiedSnapshot) And LookUp('[BETA] MONITOR KLIEN', ID = SharePointIntegration.Selected.ID).Modified <> varLastModifiedSnapshot,
    Notify("Data ini sudah diubah oleh orang lain sejak Anda membuka form. Tutup dan buka ulang item ini untuk melihat versi terbaru sebelum menyimpan.", NotificationType.Error, 6000),

    // ==========================================================================
    // 1. CEK APAKAH SCREEN 2 (TAKA) SUDAH PERNAH DIBUKA
    // ==========================================================================
    If(
        varSudahBukaTAKA <> true,
        Notify("Data pendukung di Screen 2 belum terverifikasi. Harap lakukan update/perekaman data terlebih dahulu.", NotificationType.Warning, 3000),

        // ======================================================================
        // 2. SIMPAN FORM INDUK
        // ======================================================================
        With(
            {
                TotalLangkah: CountRows(colGridUpdates),
                JumlahLangkahSelesai: CountRows(Filter(colGridUpdates, STATUS.Value = "SELESAI"))
            },
            Set(varPersenProgress, If(TotalLangkah > 0, Round((JumlahLangkahSelesai / TotalLangkah), 2), 0))
        );

        SubmitForm(SharePointForm1);

        // Cek dulu apakah submit form induk berhasil, SEBELUM lanjut ke TAKA
        If(
            !IsBlank(SharePointForm1.Error),
            Notify("Gagal menyimpan Form Induk: " & SharePointForm1.Error, NotificationType.Error, 5000),

            With(
                { CurrentID: Text(SharePointForm1.LastSubmit.ID) },

                If(
                    IsEmpty(Filter('MONITOR TAKA', 'IDMK (IDMK0)' = CurrentID)),

                    // ==============================================================
                    // SITUASI 1: DATA BARU (BATCH CREATE)
                    // ==============================================================
                    IfError(
                        Patch(
                            'MONITOR TAKA',
                            ForAll(colGridUpdates, Defaults('MONITOR TAKA')),
                            ForAll(
                                colGridUpdates As BarisBaru,
                                {
                                    'ID MK': {
                                        '@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference",
                                        Id: Value(CurrentID),
                                        Value: CurrentID
                                    },
                                    'IDMK (IDMK0)': CurrentID,
                                    NO: BarisBaru.NO,
                                    TAKA: BarisBaru.TAKA,
                                    REPORT: BarisBaru.REPORT,
                                    'TGL MULAI': BarisBaru.'TGL MULAI',
                                    'TGL SELESAI': BarisBaru.'TGL SELESAI',
                                    STATUS: If(IsBlank(BarisBaru.STATUS.Value), Blank(), {'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference", Value: BarisBaru.STATUS.Value}),
                                    PIC: If(IsBlank(BarisBaru.PIC.Claims), Blank(), {'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser", Claims: BarisBaru.PIC.Claims, Department: "", DisplayName: BarisBaru.PIC.DisplayName, Email: BarisBaru.PIC.Email, JobTitle: "", Picture: ""})
                                }
                            )
                        ),
                        Notify("Gagal menyimpan data TAKA baru: " & FirstError.Message, NotificationType.Error, 5000),
                        Notify("Data Baru Berhasil Disimpan!", NotificationType.Success, 1000)
                    ),

                    // ==============================================================
                    // SITUASI 2: MODE EDIT (BATCH UPDATE)
                    // ==============================================================
                    IfError(
                        Patch(
                            'MONITOR TAKA',
                            ForAll(
                                colGridUpdates As BarisUpdate,
                                With(
                                    { TargetRecord: LookUp('MONITOR TAKA', 'IDMK (IDMK0)' = CurrentID And NO = BarisUpdate.NO) },
                                    {
                                        ID: TargetRecord.ID,
                                        REPORT: BarisUpdate.REPORT,
                                        'TGL MULAI': BarisUpdate.'TGL MULAI',
                                        'TGL SELESAI': BarisUpdate.'TGL SELESAI',
                                        STATUS: If(IsBlank(BarisUpdate.STATUS.Value), Blank(), {'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedReference", Value: BarisUpdate.STATUS.Value}),
                                        PIC: If(IsBlank(BarisUpdate.PIC.Claims), Blank(), {'@odata.type': "#Microsoft.Azure.Connectors.SharePoint.SPListExpandedUser", Claims: BarisUpdate.PIC.Claims, Department: "", DisplayName: BarisUpdate.PIC.DisplayName, Email: BarisUpdate.PIC.Email, JobTitle: "", Picture: ""})
                                    }
                                )
                            )
                        ),
                        Notify("Gagal menyimpan update TAKA: " & FirstError.Message, NotificationType.Error, 5000),
                        Notify("Data Berhasil Disimpan!", NotificationType.Success, 1000)
                    )
                );

                Navigate(FormScreen1);
                Set(varIsSaving, false);
                RequestHide()
            )
        )
    )
)
