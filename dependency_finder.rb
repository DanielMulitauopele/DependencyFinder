require 'write_xlsx'
require 'find'
require 'pry'

class DependencyFinder
    def initialize
        @commands = {
            0 => {header: 'API Name'},
            1 => {header: 'Apex Class', file_type: '/Users/daniel.m/Projects/Polaris/**/*.cls'},
            2 => {header: 'Approval Process', file_type: '/Users/daniel.m/Projects/Polaris/**/**.approvalProcess-meta.xml'},
            3 => {header: 'Custom Metadata', file_type: '/Users/daniel.m/Projects/Polaris/**/**.md-meta.xml'},
            4 => {header: 'Field', file_type: '/Users/daniel.m/Projects/Polaris/**/**.field-meta.xml'},
            5 => {header: 'Flow', file_type: '/Users/daniel.m/Projects/Polaris/**/**.flow-meta.xml'},
            6 => {header: 'Global Value Set', file_type: '/Users/daniel.m/Projects/Polaris/**/**.globalValueSet-meta.xml'},
            7 => {header: 'Layout', file_type: '/Users/daniel.m/Projects/Polaris/**/**.layout-meta.xml'},
            8 => {header: 'Page', file_type: '/Users/daniel.m/Projects/Polaris/**/**.page-meta.xml'},
            9 => {header: 'Quick Action', file_type: '/Users/daniel.m/Projects/Polaris/**/**.quickAction-meta.xml'},
            10 => {header: 'Report', file_type: '/Users/daniel.m/Projects/Polaris/**/**.report-meta.xml'},
            11 => {header: 'Standard Value Set', file_type: '/Users/daniel.m/Projects/Polaris/**/**.standardValueSet.xml'},
            12 => {header: 'Trigger', file_type: '/Users/daniel.m/Projects/Polaris/**/**.Trigger'},
            13 => {header: 'Validation Rule', file_type: '/Users/daniel.m/Projects/Polaris/**/**.validationRule-meta.xml'},
            14 => {header: 'Workflow', file_type: '/Users/daniel.m/Projects/Polaris/**/**.workflow-meta.xml'}
        }

        @font = {
            :font  => 'Calibri',
            :size  => 14,
            :color => 'black',
            :bold  => 1
        }
    end

    def write_to_excel(file, inputs)
        f = File.new(file, chomp: true)
        workbook = WriteXLSX.new('Opportunity_Field_Dependencies.xlsx')
        worksheet = workbook.add_worksheet
        
        # write_headers(workbook, worksheet, inputs)
        write_lines(f, workbook, worksheet, inputs)

        workbook.close
        sleep 2
        puts 'Excel file created!'
    end

    def write_lines(file, workbook, worksheet, inputs)
        header_format = workbook.add_format(@font)
        inputs.each_with_index do |value, index|
            worksheet.write(0, index, @commands[value][:header], header_format) 
        end 

        row = 1
        file.readlines.each do |line|
            inputs.each_with_index do |value, index|
                if index == 0
                    worksheet.write(row, index, line)
                else
                    worksheet.write(row, index, dependency_search(line, @commands[value][:file_type]))
                end 
            end

            output = row.to_s + ' fields reviewed...' + "\r"
            print output
            $stdout.flush
            row += 1
        end
    end

    def dependency_search(field, dependency_type)
        dependencies = []

        Dir.glob(dependency_type) do |file|
            File.readlines(file).each do |line|
                dependencies << file if line.include?(field.chomp)
            end
        end

        if dependencies.empty?
            return 'no results'
        else 
            return dependencies.uniq.join(', ')
        end
    end
end 