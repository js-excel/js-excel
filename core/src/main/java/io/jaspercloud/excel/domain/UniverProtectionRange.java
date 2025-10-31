package io.jaspercloud.excel.domain;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

@AllArgsConstructor
@NoArgsConstructor
@Setter
@Getter
public class UniverProtectionRange {

    private String range;
    private Boolean locked;

    @Override
    public String toString() {
        return "UniverProtectionRange{" +
                "range='" + range + '\'' +
                ", locked=" + locked +
                '}';
    }
}
