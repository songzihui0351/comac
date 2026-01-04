package main;

import lombok.Data;

import java.util.Objects;


@Data
public class Bill {
    private String name;
    private String dept;
    private String arrival;
    private String departure;
    private int days;

    public Bill(String name, String dept, String date) {
        this.name = name;
        this.dept = dept;
        this.arrival = formater(date);
        this.departure = formater(date);
        this.days = 1;
    }

    public void incrementDays(String date, String remark) {
        this.departure = formater(date);
        if (remark.equals("离开")) return;
        this.days++;
    }

    private String formater(String date) {
        return "2025/" + date.replaceAll("\\.", "/");
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (!(o instanceof Bill bill)) return false;
        return Objects.equals(this.name, bill.name)
                && Objects.equals(this.dept, bill.dept);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, dept);
    }
}
