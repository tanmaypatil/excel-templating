package com.excelgen;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Test for Phone functionality in Person class
 */
class PhoneTest {

    @Test
    void testPersonWithoutPhones() {
        Person person = new Person("John Doe", 25, "Jane Doe");

        assertNotNull(person.getPhones(), "Phone list should be initialized");
        assertTrue(person.getPhones().isEmpty(), "Phone list should be empty initially");
        assertFalse(person.hasPhones(), "Should return false when no phones");

        System.out.println("✓ Person created without phones");
    }

    @Test
    void testAddPhoneWithObject() {
        Person person = new Person("Alice Smith", 30, "Bob Smith");

        Phone mobilePhone = new Phone("Mobile", "+1-555-1234");
        person.addPhone(mobilePhone);

        assertEquals(1, person.getPhones().size(), "Should have 1 phone");
        assertTrue(person.hasPhones(), "Should return true when phones exist");
        assertEquals("Mobile", person.getPhones().get(0).getPhoneType());
        assertEquals("+1-555-1234", person.getPhones().get(0).getPhoneNo());

        System.out.println("✓ Phone added using Phone object");
    }

    @Test
    void testAddPhoneWithParameters() {
        Person person = new Person("Tom Brown", 28, "Sarah Brown");

        person.addPhone("Home", "+1-555-5678");
        person.addPhone("Work", "+1-555-9012");

        assertEquals(2, person.getPhones().size(), "Should have 2 phones");
        assertTrue(person.hasPhones(), "Should return true when phones exist");

        assertEquals("Home", person.getPhones().get(0).getPhoneType());
        assertEquals("+1-555-5678", person.getPhones().get(0).getPhoneNo());

        assertEquals("Work", person.getPhones().get(1).getPhoneType());
        assertEquals("+1-555-9012", person.getPhones().get(1).getPhoneNo());

        System.out.println("✓ Multiple phones added using parameters");
    }

    @Test
    void testPersonWithMultiplePhones() {
        Person person = new Person("Emma Wilson", 35, "David Wilson");

        person.addPhone("Mobile", "+1-555-1111");
        person.addPhone("Home", "+1-555-2222");
        person.addPhone("Work", "+1-555-3333");
        person.addPhone("Fax", "+1-555-4444");

        assertEquals(4, person.getPhones().size(), "Should have 4 phones");
        assertTrue(person.hasPhones(), "Should have phones");

        System.out.println("✓ Person with multiple phones:");
        for (Phone phone : person.getPhones()) {
            System.out.println("  - " + phone.getPhoneType() + ": " + phone.getPhoneNo());
        }
    }

    @Test
    void testPersonWithAddressAndPhones() {
        Person person = new Person("Mike Davis", 40, "Linda Davis");

        Address address = new Address("Home", "123 Main St");
        person.setAddress(address);

        person.addPhone("Mobile", "+1-555-7777");
        person.addPhone("Home", "+1-555-8888");

        assertTrue(person.isAddressExists(), "Should have address");
        assertTrue(person.hasPhones(), "Should have phones");
        assertEquals(2, person.getPhones().size(), "Should have 2 phones");

        System.out.println("✓ Person with both address and phones");
        System.out.println("  Address: " + address.getAddressLine());
        System.out.println("  Phones:");
        for (Phone phone : person.getPhones()) {
            System.out.println("    - " + phone.getPhoneType() + ": " + phone.getPhoneNo());
        }
    }
}
