package com.bulain.jxls.export;

import java.io.IOException;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.sql.DataSource;

import net.sf.jxls.report.ReportManager;
import net.sf.jxls.report.ReportManagerImpl;
import net.sf.jxls.report.ResultSetCollection;
import net.sf.jxls.transformer.Configuration;
import net.sf.jxls.transformer.XLSTransformer;

import org.apache.commons.beanutils.RowSetDynaClass;
import org.junit.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.ClassPathResource;

import com.bulain.common.dataset.DataSet;
import com.bulain.common.model.Person;
import com.bulain.common.pojo.PersonSearch;
import com.bulain.common.service.PersonService;
import com.bulain.common.test.ServiceTestCase;

@DataSet(file = "test-data/init_persons.xml")
public class PersonTest extends ServiceTestCase {
    @Autowired
    private PersonService personService;
    @Autowired
    private DataSource dataSource;

    @Test
    public void testMybatis() throws Exception {
        PersonSearch search = new PersonSearch();
        List<Person> listPerson = personService.find(search);

        Map<String, Object> beans = new HashMap<String, Object>();
        beans.put("person", listPerson);

        Configuration config = new Configuration();
        XLSTransformer transformer = new XLSTransformer(config);
        String templatePath = getTemplatePath("xls-template/persons.xls");
        transformer.transformXLS(templatePath, beans, "target/mybatis.xls");
    }

    @Test
    public void testRowSetDyna() throws Exception {
        Connection conn = dataSource.getConnection();

        StringBuffer query = new StringBuffer();
        query.append("select id, first_name firstName, last_name, created_by, created_at, updated_by, updated_at ");
        query.append("from persons order by id ");
        PreparedStatement ps = conn.prepareStatement(query.toString());
        ResultSet rs = ps.executeQuery();
        RowSetDynaClass rsdc = new RowSetDynaClass(rs, false);

        Map<String, Object> beans = new HashMap<String, Object>();
        beans.put("person", rsdc.getRows());

        Configuration config = new Configuration();
        XLSTransformer transformer = new XLSTransformer(config);
        String templatePath = getTemplatePath("xls-template/rowSetDyna.xls");
        transformer.transformXLS(templatePath, beans, "target/rowSetDyna.xls");

        rs.close();
        ps.close();
        conn.close();
    }

    @Test
    public void testReporting() throws Exception {
        Connection conn = dataSource.getConnection();

        Map<String, Object> beans = new HashMap<String, Object>();
        ReportManager reportManager = new ReportManagerImpl(conn, beans);
        beans.put("rm", reportManager);

        Configuration config = new Configuration();
        XLSTransformer transformer = new XLSTransformer(config);
        String templatePath = getTemplatePath("xls-template/reporting.xls");
        transformer.transformXLS(templatePath, beans, "target/reporting.xls");

        conn.close();
    }
    
    @Test
    public void testResultSetCollection() throws Exception{
        Connection conn = dataSource.getConnection();

        StringBuffer query = new StringBuffer();
        query.append("select id, first_name firstName, last_name, created_by, created_at, updated_by, updated_at ");
        query.append("from persons order by id ");
        PreparedStatement ps = conn.prepareStatement(query.toString());
        ResultSet rs = ps.executeQuery();
        ResultSetCollection rsc = new ResultSetCollection(rs, 100, true);

        Map<String, Object> beans = new HashMap<String, Object>();
        beans.put("person", rsc);

        Configuration config = new Configuration();
        XLSTransformer transformer = new XLSTransformer(config);
        String templatePath = getTemplatePath("xls-template/rowSetDyna.xls");
        transformer.transformXLS(templatePath, beans, "target/resultSetCollection.xls");

        rs.close();
        ps.close();
        conn.close();
    }

    private String getTemplatePath(String template) throws IOException {
        ClassPathResource resource = new ClassPathResource(template);
        return resource.getFile().getAbsolutePath();
    }

}
