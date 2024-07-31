package com.example.demo.users_mng.pageable;

import com.fasterxml.jackson.core.JsonGenerator;
import com.fasterxml.jackson.databind.SerializerProvider;
import com.fasterxml.jackson.databind.ser.std.StdSerializer;

import java.io.IOException;

@SuppressWarnings("all")
public class SimplePageSerializer extends StdSerializer<SimplePage> {

    public SimplePageSerializer() {
        this(null);
    }

    public SimplePageSerializer(Class<SimplePage> t) {
        super(t);
    }

    @Override
    public void serialize(SimplePage simplePage, JsonGenerator gen, SerializerProvider provider) throws IOException {
        gen.writeStartObject();
        gen.writeObjectField("content", simplePage.getContent());
        gen.writeNumberField("totalElements", simplePage.getTotalElements());
        gen.writeNumberField("totalPages", simplePage.getTotalPages());
        gen.writeNumberField("page", simplePage.getPage());
        gen.writeNumberField("size", simplePage.getSize());
        gen.writeObjectField("sort", simplePage.getSortList());
        gen.writeEndObject();
    }
}
