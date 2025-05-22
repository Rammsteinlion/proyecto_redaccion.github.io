const filteredList = computed(() => {
    const activeList: any[] = filteredPromotionsByState.value.active || [];
    const expiredList: any[] = filteredPromotionsByState.value.expired || [];

    // Función para normalizar el tipo (convertir string a número si es necesario)
    const getNormalizedType = (promo: any): number | undefined => {
        if (promo.type === undefined || promo.type === null) return undefined;
        return Number(promo.type); // Convierte "1" → 1, "2" → 2, etc.
    };

    if (selectedType.value !== 0) {
        const selectedTypeList =
            selectedType.value === 1
                ? activeList.filter((promo) => getNormalizedType(promo) === 1)
                : expiredList.filter((promo) => getNormalizedType(promo) === 1);
        
        const selectedCasinoList =
            selectedType.value === 2
                ? activeList.filter((promo) => getNormalizedType(promo) === 2)
                : expiredList.filter((promo) => getNormalizedType(promo) === 2);

        if (selectedStatus.value === 1) {
            return selectedType.value === 1 ? selectedTypeList : selectedCasinoList;
        } else {
            return selectedType.value === 1
                ? expiredList.filter((promo) => getNormalizedType(promo) === 1)
                : expiredList.filter((promo) => getNormalizedType(promo) === 2);
        }
    } else {
        return selectedStatus.value === 1 ? activeList : expiredList;
    }
});